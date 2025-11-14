from dataclasses import dataclass
from datetime import datetime, timedelta
from decimal import Decimal
from typing import List, Dict, Optional
from enum import Enum

from dash import dcc, no_update, Output, Input, State

from supply_chain_analytics import app


class StatutCommande(Enum):
    EN_ATTENTE = "en_attente"
    VALIDEE = "validee"
    ENVOYEE = "envoyee"
    RECUE = "recue"
    ANNULEE = "annulee"


@dataclass
class Produit:
    id: str
    nom: str
    reference: str
    stock_actuel: int
    stock_minimum: int  # Seuil d'alerte
    stock_securite: int  # Stock de sécurité
    unite: str  # "pièce", "kg", "litre"
    prix_achat_unitaire: Decimal
    fournisseur_id: str
    delai_livraison_jours: int
    quantite_commande_minimum: int  # MOQ (Minimum Order Quantity)
    conditionnement: int  # Par combien commander (ex: par carton de 12)

    def est_en_rupture(self) -> bool:
        """Produit actuellement en rupture"""
        return self.stock_actuel <= 0

    def est_critique(self) -> bool:
        """Stock en dessous du seuil minimum"""
        return self.stock_actuel <= self.stock_minimum

    def jours_avant_rupture(self, ventes_moyennes_jour: float) -> float:
        """Estimation des jours restants avant rupture"""
        if ventes_moyennes_jour <= 0:
            return float('inf')
        return self.stock_actuel / ventes_moyennes_jour

    def quantite_recommandee(self, ventes_moyennes_jour: float,
                             horizon_jours: int = 30) -> int:
        """
        Calcul intelligent de la quantité à commander
        Prend en compte :
        - Ventes moyennes
        - Délai de livraison
        - Stock de sécurité
        - MOQ et conditionnement
        """
        # Besoin = (ventes pendant délai + horizon) - stock actuel + stock sécu
        besoin_brut = (ventes_moyennes_jour * (self.delai_livraison_jours + horizon_jours)
                       - self.stock_actuel + self.stock_securite)

        # Arrondi au conditionnement supérieur
        if besoin_brut < self.quantite_commande_minimum:
            besoin_brut = self.quantite_commande_minimum

        quantite = int(besoin_brut)

        # Arrondir au conditionnement supérieur
        if self.conditionnement > 1:
            reste = quantite % self.conditionnement
            if reste > 0:
                quantite += (self.conditionnement - reste)

        return max(quantite, self.quantite_commande_minimum)


@dataclass
class Fournisseur:
    id: str
    nom: str
    email: str
    telephone: str
    adresse: str
    delai_paiement_jours: int  # 30, 60, 90 jours
    conditions_particulieres: Optional[str]
    montant_minimum_commande: Decimal  # Montant minimum de commande


@dataclass
class LigneCommande:
    produit: Produit
    quantite: int
    prix_unitaire: Decimal
    remise_pourcent: Decimal = Decimal("0")

    @property
    def montant_ht(self) -> Decimal:
        base = self.prix_unitaire * Decimal(str(self.quantite))
        return base * (Decimal("1") - self.remise_pourcent / Decimal("100"))


@dataclass
class BonCommande:
    id: str
    numero: str  # Format : BC-2025-001
    date_creation: datetime
    fournisseur: Fournisseur
    lignes: List[LigneCommande]
    statut: StatutCommande
    commentaire: Optional[str]
    date_livraison_souhaitee: Optional[datetime]

    @property
    def montant_total_ht(self) -> Decimal:
        return sum(ligne.montant_ht for ligne in self.lignes)

    @property
    def montant_tva(self) -> Decimal:
        return self.montant_total_ht * Decimal("0.20")  # TVA 20%

    @property
    def montant_total_ttc(self) -> Decimal:
        return self.montant_total_ht + self.montant_tva


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


from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from decimal import Decimal
from io import BytesIO


class GenerateurBonCommandeAvecRemises:
    """
    Générateur de bons de commande PDF avec calculs automatiques
    Prend en compte remises et escomptes
    """

    def __init__(self, entreprise_info: dict):
        self.entreprise = entreprise_info
        self.TVA_TAUX = Decimal("0.18")

    def calculer_ligne_complete(self, quantite: int, prix_unitaire: Decimal,
                                remise_pourcent: Decimal, escompte_pourcent: Decimal) -> dict:
        """
        Calcul complet d'une ligne avec remise et escompte

        Formule:
        1. Montant brut = Qté × PU
        2. Montant après remise = Montant brut × (1 - Remise/100)
        3. Montant HT = Montant après remise × (1 - Escompte/100)
        4. TVA = Montant HT × 18%
        5. TTC = Montant HT + TVA
        """
        montant_brut = Decimal(str(quantite)) * prix_unitaire

        # Application remise
        montant_apres_remise = montant_brut * (Decimal("1") - remise_pourcent / Decimal("100"))

        # Application escompte
        montant_ht = montant_apres_remise * (Decimal("1") - escompte_pourcent / Decimal("100"))

        # TVA et TTC
        montant_tva = montant_ht * self.TVA_TAUX
        montant_ttc = montant_ht + montant_tva

        return {
            'montant_brut': montant_brut,
            'montant_apres_remise': montant_apres_remise,
            'montant_ht': montant_ht,
            'montant_tva': montant_tva,
            'montant_ttc': montant_ttc,
            'economie_remise': montant_brut - montant_apres_remise,
            'economie_escompte': montant_apres_remise - montant_ht
        }

    def generer_pdf(self, bon_commande: dict) -> bytes:
        """
        Génère le PDF du bon de commande

        bon_commande structure:
        {
            'numero': 'BC-2025-001',
            'date': '2025-01-15',
            'fournisseur': {...},
            'lignes': [
                {
                    'reference': 'RIZ-001',
                    'designation': 'Riz parfumé 25kg',
                    'quantite': 100,
                    'unite': 'sac',
                    'prix_unitaire': 15000.0,
                    'remise': 5.0,  # Pourcentage
                    'escompte': 2.0  # Pourcentage
                },
                ...
            ],
            'date_livraison_souhaitee': '2025-01-22',
            'conditions_paiement': '30 jours fin de mois'
        }
        """

        buffer = BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=2 * cm, leftMargin=2 * cm)
        elements = []
        styles = getSampleStyleSheet()

        # En-tête
        elements.append(Paragraph(f"<b>BON DE COMMANDE N° {bon_commande['numero']}</b>", styles['Title']))
        elements.append(Spacer(1, 0.5 * cm))

        # Informations entreprise
        info_entreprise = f"""
        <b>{self.entreprise['nom']}</b><br/>
        {self.entreprise['adresse']}<br/>
        Tél: {self.entreprise['telephone']}<br/>
        Email: {self.entreprise['email']}
        """
        elements.append(Paragraph(info_entreprise, styles['Normal']))
        elements.append(Spacer(1, 1 * cm))

        # Informations commande
        data_info = [
            ['Date:', bon_commande['date']],
            ['Date livraison souhaitée:', bon_commande.get('date_livraison_souhaitee', 'À définir')],
            ['Fournisseur:', bon_commande['fournisseur']['nom']],
            ['Email:', bon_commande['fournisseur']['email']]
        ]
        table_info = Table(data_info, colWidths=[5 * cm, 10 * cm])
        table_info.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (0, -1), colors.lightgrey),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey)
        ]))
        elements.append(table_info)
        elements.append(Spacer(1, 1 * cm))

        # Tableau des produits avec calculs
        data_lignes = [[
            'Réf.',
            'Désignation',
            'Qté',
            'PU HT',
            'Remise %',
            'Escompte %',
            'Total HT',
            'TVA 18%',
            'Total TTC'
        ]]

        total_brut = Decimal("0")
        total_ht = Decimal("0")
        total_tva = Decimal("0")
        total_ttc = Decimal("0")
        total_remises = Decimal("0")
        total_escomptes = Decimal("0")

        for ligne in bon_commande['lignes']:
            calculs = self.calculer_ligne_complete(
                quantite=ligne['quantite'],
                prix_unitaire=Decimal(str(ligne['prix_unitaire'])),
                remise_pourcent=Decimal(str(ligne.get('remise', 0))),
                escompte_pourcent=Decimal(str(ligne.get('escompte', 0)))
            )

            total_brut += calculs['montant_brut']
            total_ht += calculs['montant_ht']
            total_tva += calculs['montant_tva']
            total_ttc += calculs['montant_ttc']
            total_remises += calculs['economie_remise']
            total_escomptes += calculs['economie_escompte']

            data_lignes.append([
                ligne['reference'],
                ligne['designation'][:40],
                str(ligne['quantite']),
                f"{float(ligne['prix_unitaire']):,.0f}",
                f"{ligne.get('remise', 0):.1f}" if ligne.get('remise', 0) > 0 else "-",
                f"{ligne.get('escompte', 0):.1f}" if ligne.get('escompte', 0) > 0 else "-",
                f"{float(calculs['montant_ht']):,.0f}",
                f"{float(calculs['montant_tva']):,.0f}",
                f"{float(calculs['montant_ttc']):,.0f}"
            ])

        # Ligne vide avant totaux
        data_lignes.append(['', '', '', '', '', '', '', '', ''])

        # Totaux avec économies
        if total_remises > 0 or total_escomptes > 0:
            data_lignes.append(['', '', '', '', '', 'Montant brut:', f"{float(total_brut):,.0f}", '', ''])
            if total_remises > 0:
                data_lignes.append(['', '', '', '', '', 'Remises totales:', f"-{float(total_remises):,.0f}", '', ''])
            if total_escomptes > 0:
                data_lignes.append(['', '', '', '', '', 'Escomptes totaux:', f"-{float(total_escomptes):,.0f}", '', ''])

        data_lignes.append(['', '', '', '', '', 'TOTAL HT:', f"{float(total_ht):,.0f}", '', ''])
        data_lignes.append(['', '', '', '', '', 'TVA (18%):', f"{float(total_tva):,.0f}", '', ''])
        data_lignes.append(['', '', '', '', '', 'TOTAL TTC:', '', '', f"{float(total_ttc):,.0f}"])

        # Création table
        col_widths = [2 * cm, 5 * cm, 1.5 * cm, 2 * cm, 1.5 * cm, 1.5 * cm, 2.5 * cm, 2 * cm, 2.5 * cm]
        table_lignes = Table(data_lignes, colWidths=col_widths)

        # Style du tableau
        style_list = [
            # En-tête
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2C3E50')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 9),
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),

            # Corps
            ('GRID', (0, 0), (-1, -8), 0.5, colors.grey),
            ('FONTSIZE', (0, 1), (-1, -1), 8),
            ('ALIGN', (2, 1), (2, -8), 'CENTER'),  # Quantité
            ('ALIGN', (3, 1), (-1, -8), 'RIGHT'),  # Prix et montants

            # Colonnes remise et escompte
            ('BACKGROUND', (4, 1), (4, -8), colors.lightyellow),
            ('BACKGROUND', (5, 1), (5, -8), colors.HexColor('#FFE5CC')),

            # Totaux
            ('FONTNAME', (5, -7), (-1, -1), 'Helvetica-Bold'),
            ('ALIGN', (5, -7), (-1, -1), 'RIGHT'),

            # Total TTC
            ('BACKGROUND', (5, -1), (-1, -1), colors.HexColor('#27AE60')),
            ('TEXTCOLOR', (5, -1), (-1, -1), colors.whitesmoke),
            ('FONTSIZE', (5, -1), (-1, -1), 11),
        ]

        table_lignes.setStyle(TableStyle(style_list))
        elements.append(table_lignes)
        elements.append(Spacer(1, 1 * cm))

        # Économies réalisées (si applicable)
        if total_remises + total_escomptes > 0:
            elements.append(Paragraph("<b>💰 Économies réalisées</b>", styles['Heading3']))
            economie_text = f"""
            Grâce aux remises et escomptes appliqués, vous économisez <b>{float(total_remises + total_escomptes):,.0f} FCFA</b><br/>
            - Remises: {float(total_remises):,.0f} FCFA<br/>
            - Escomptes: {float(total_escomptes):,.0f} FCFA
            """
            elements.append(Paragraph(economie_text, styles['Normal']))
            elements.append(Spacer(1, 0.5 * cm))

        # Conditions
        elements.append(Paragraph("<b>Conditions de paiement:</b>", styles['Heading3']))
        elements.append(Paragraph(bon_commande.get('conditions_paiement', 'Selon accord'), styles['Normal']))
        elements.append(Spacer(1, 0.5 * cm))

        # Notes explicatives
        elements.append(Paragraph("<b>Notes importantes:</b>", styles['Heading3']))
        notes = """
        • <b>Remise</b>: Réduction commerciale appliquée sur le prix unitaire<br/>
        • <b>Escompte</b>: Réduction financière pour paiement anticipé (appliquée après remise)<br/>
        • Formule: TTC = (PU × Qté × (1 - Remise/100) × (1 - Escompte/100)) × 1.18
        """
        elements.append(Paragraph(notes, styles['Normal']))

        # Signature
        elements.append(Spacer(1, 2 * cm))
        elements.append(Paragraph("""
        <para align=right>
        <b>Signature et cachet</b><br/><br/><br/>
        _______________________<br/>
        {nom}
        </para>
        """.format(nom=self.entreprise['nom']), styles['Normal']))

        # Construction
        doc.build(elements)
        pdf_bytes = buffer.getvalue()
        buffer.close()

        return pdf_bytes


# ============================================
# UTILISATION DANS VOTRE CALLBACK DASH
# ============================================

@app.callback(
    Output("download-po-pdf", "data"),
    Input("btn-generer-pdf-final", "n_clicks"),
    State("bon-commande-data-store", "data"),  # Store contenant les données éditées
    prevent_initial_call=True
)
def generer_pdf_avec_remises(n_clicks, bon_commande_data):
    """
    Callback Dash pour générer le PDF après édition
    """
    if not n_clicks or not bon_commande_data:
        return no_update

    try:
        entreprise_info = {
            'nom': 'MAAD DISTRIBUTION',
            'adresse': 'Rue 123, Dakar, Sénégal',
            'telephone': '+221 77 105 16 65',
            'email': 'tony.sarre@maad.io'
        }

        generateur = GenerateurBonCommandeAvecRemises(entreprise_info)
        pdf_bytes = generateur.generer_pdf(bon_commande_data)

        filename = f"BC_{bon_commande_data['numero']}.pdf"

        print(f"✅ PDF généré : {filename}")
        print(f"   Lignes : {len(bon_commande_data['lignes'])}")

        # Calcul du total pour log
        total_ttc = sum(
            generateur.calculer_ligne_complete(
                ligne['quantite'],
                Decimal(str(ligne['prix_unitaire'])),
                Decimal(str(ligne.get('remise', 0))),
                Decimal(str(ligne.get('escompte', 0)))
            )['montant_ttc']
            for ligne in bon_commande_data['lignes']
        )

        print(f"   Total TTC : {float(total_ttc):,.0f} FCFA")

        return dcc.send_bytes(pdf_bytes, filename=filename)

    except Exception as e:
        print(f"❌ Erreur génération PDF : {e}")
        return no_update