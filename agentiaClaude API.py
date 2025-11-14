import anthropic
import json
from typing import List, Dict


class AgentAchatIA:
    """
    Agent IA qui analyse le stock et recommande des commandes intelligentes
    Utilise Claude API pour des décisions contextuelles
    """

    def __init__(self, api_key: str = None):
        # L'API key est gérée automatiquement dans les artifacts
        self.client = anthropic.Anthropic()

    def analyser_stock_et_recommander(self, produits: List[Produit],
                                      historique_ventes: Dict[str, List[float]]) -> Dict:
        """
        Analyse intelligente des produits en rupture/critique
        et génération de recommandations d'achat

        Args:
            produits: Liste des produits à analyser
            historique_ventes: {produit_id: [ventes_j1, ventes_j2, ...]}

        Returns:
            Recommandations structurées par fournisseur
        """

        # Préparation des données pour l'IA
        donnees_produits = []
        for produit in produits:
            ventes_historique = historique_ventes.get(produit.id, [])
            ventes_moy = sum(ventes_historique[-30:]) / len(ventes_historique[-30:]) if ventes_historique else 0

            jours_restants = produit.jours_avant_rupture(ventes_moy)

            donnees_produits.append({
                'id': produit.id,
                'nom': produit.nom,
                'reference': produit.reference,
                'stock_actuel': produit.stock_actuel,
                'stock_minimum': produit.stock_minimum,
                'ventes_moyennes_jour': round(ventes_moy, 2),
                'jours_avant_rupture': round(jours_restants, 1) if jours_restants != float('inf') else 'infini',
                'fournisseur_id': produit.fournisseur_id,
                'delai_livraison': produit.delai_livraison_jours,
                'prix_unitaire': float(produit.prix_achat_unitaire),
                'est_en_rupture': produit.est_en_rupture(),
                'est_critique': produit.est_critique()
            })

        # Prompt pour Claude
        prompt = f"""Tu es un expert en gestion de supply chain. Analyse les données suivantes et recommande les actions d'achat prioritaires.

DONNÉES DES PRODUITS :
{json.dumps(donnees_produits, indent=2, ensure_ascii=False)}

CONSIGNES :
1. Identifie les produits nécessitant une commande URGENTE (rupture ou <7 jours restants)
2. Identifie les produits nécessitant une commande NORMALE (7-14 jours restants)
3. Pour chaque produit, calcule la quantité optimale à commander
4. Regroupe les commandes par fournisseur
5. Priorise selon l'impact business (risque de perte de vente)

Réponds UNIQUEMENT en JSON avec cette structure exacte :
{{
  "analyse_globale": "texte libre d'analyse",
  "commandes_urgentes": [
    {{
      "fournisseur_id": "F001",
      "priorite": "URGENTE",
      "produits": [
        {{
          "produit_id": "P001",
          "quantite_recommandee": 100,
          "justification": "Rupture imminente sous 3 jours"
        }}
      ],
      "montant_total_estime": 1500.00
    }}
  ],
  "commandes_normales": [...]
}}

IMPORTANT : Ne retourne QUE le JSON, sans texte avant ou après, sans balises markdown.
"""

        try:
            # Appel à Claude API
            response = self.client.messages.create(
                model="claude-sonnet-4-20250514",
                max_tokens=4000,
                messages=[{
                    "role": "user",
                    "content": prompt
                }]
            )

            # Extraction du texte
            texte_reponse = response.content[0].text

            # Nettoyage (au cas où Claude ajoute des backticks)
            texte_reponse = texte_reponse.replace('```json', '').replace('```', '').strip()

            # Parse JSON
            recommandations = json.loads(texte_reponse)

            return recommandations

        except json.JSONDecodeError as e:
            print(f"Erreur de parsing JSON : {e}")
            print(f"Réponse brute : {texte_reponse}")
            raise Exception("L'IA n'a pas retourné un JSON valide")

        except Exception as e:
            print(f"Erreur lors de l'appel à l'IA : {e}")
            raise

    def generer_plan_achat_complet(self, produits: List[Produit],
                                   fournisseurs: Dict[str, Fournisseur],
                                   historique_ventes: Dict[str, List[float]],
                                   budget_max: Optional[Decimal] = None) -> List[BonCommande]:
        """
        Génère automatiquement les bons de commande à partir des recommandations IA
        """

        # Étape 1 : Analyse IA
        recommandations = self.analyser_stock_et_recommander(produits, historique_ventes)

        print("📊 ANALYSE IA TERMINÉE")
        print(f"Analyse globale : {recommandations.get('analyse_globale', 'N/A')}")
        print(f"Commandes urgentes : {len(recommandations.get('commandes_urgentes', []))}")
        print(f"Commandes normales : {len(recommandations.get('commandes_normales', []))}")

        # Étape 2 : Conversion en bons de commande
        bons_commande = []

        toutes_commandes = (recommandations.get('commandes_urgentes', []) +
                            recommandations.get('commandes_normales', []))

        for index, cmd_recommandee in enumerate(toutes_commandes):
            fournisseur_id = cmd_recommandee['fournisseur_id']
            fournisseur = fournisseurs.get(fournisseur_id)

            if not fournisseur:
                print(f"⚠️ Fournisseur {fournisseur_id} introuvable, commande ignorée")
                continue

            # Création des lignes de commande
            lignes = []
            for prod_rec in cmd_recommandee['produits']:
                produit = next((p for p in produits if p.id == prod_rec['produit_id']), None)
                if not produit:
                    continue

                ligne = LigneCommande(
                    produit=produit,
                    quantite=prod_rec['quantite_recommandee'],
                    prix_unitaire=produit.prix_achat_unitaire
                )
                lignes.append(ligne)

            if not lignes:
                continue

            # Vérification montant minimum commande
            montant_total = sum(l.montant_ht for l in lignes)
            if montant_total < fournisseur.montant_minimum_commande:
                print(
                    f"⚠️ Montant {montant_total}€ < minimum {fournisseur.montant_minimum_commande}€ pour {fournisseur.nom}")
                # Option : ajuster automatiquement ou alerter

            # Vérification budget
            if budget_max and montant_total > budget_max:
                print(f"⚠️ Budget dépassé : {montant_total}€ > {budget_max}€")
                # Option : prioriser ou alerter

            # Création du bon de commande
            bon = BonCommande(
                id=f"BC-{datetime.now().strftime('%Y%m%d')}-{index + 1:03d}",
                numero=f"BC-{datetime.now().year}-{index + 1:03d}",
                date_creation=datetime.now(),
                fournisseur=fournisseur,
                lignes=lignes,
                statut=StatutCommande.EN_ATTENTE,
                commentaire=cmd_recommandee.get('justification'),
                date_livraison_souhaitee=datetime.now() + timedelta(days=7)
            )

            bons_commande.append(bon)

        return bons_commande