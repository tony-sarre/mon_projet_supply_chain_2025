// ========================================
// 🎨 COLORISATION TABLEAU DASH
// ========================================
/*console.log("🎨 Script colorize.js chargé");

function colorizeTableRows() {
    console.log("🔍 colorizeTableRows() appelée");

    const table = document.querySelector('.dash-table-container table');

    if (!table) {
        console.warn("⚠️ Table non trouvée");
        return;
    }

    console.log("✅ Table trouvée :", table);

    const rows = table.querySelectorAll('tbody tr');
    console.log(`📊 ${rows.length} lignes détectées`);

    let coloredCount = 0;

    rows.forEach((row, rowIndex) => {
        const cells = row.querySelectorAll('td');

        cells.forEach((cell, cellIndex) => {
            const text = cell.textContent.trim();

            // 🔴 ORDER NOW
            if (text === 'ORDER NOW') {
                console.log(`🔴 ORDER NOW trouvé - ligne ${rowIndex}`);

                // Ligne entière
                row.style.backgroundColor = 'rgba(239, 68, 68, 0.15)';
                row.style.transition = 'all 0.3s ease';

                cells.forEach(c => {
                    c.style.color = '#fecaca';
                    c.style.fontWeight = '600';
                });

                // Badge
                cell.style.backgroundColor = 'rgba(239, 68, 68, 0.5)';
                cell.style.color = '#ffffff';
                cell.style.fontWeight = '800';
                cell.style.border = '3px solid #ef4444';
                cell.style.borderRadius = '8px';
                cell.style.padding = '8px';
                cell.style.textTransform = 'uppercase';

                // Bordure gauche sur product_name (2ème colonne)
                if (cells[1]) {
                    cells[1].style.borderLeft = '5px solid #ef4444';
                    cells[1].style.backgroundColor = 'rgba(239, 68, 68, 0.25)';
                    cells[1].style.color = '#fee2e2';
                    cells[1].style.fontWeight = '700';
                }

                coloredCount++;
            }

            // 🟠 ORDER NOT URGENT
            else if (text === 'ORDER NOT URGENT') {
                console.log(`🟠 ORDER NOT URGENT trouvé - ligne ${rowIndex}`);

                row.style.backgroundColor = 'rgba(245, 158, 11, 0.12)';

                cells.forEach(c => {
                    c.style.color = '#fde68a';
                    c.style.fontWeight = '500';
                });

                cell.style.backgroundColor = 'rgba(245, 158, 11, 0.5)';
                cell.style.color = '#ffffff';
                cell.style.fontWeight = '700';
                cell.style.border = '3px solid #f59e0b';
                cell.style.borderRadius = '8px';
                cell.style.padding = '8px';
                cell.style.textTransform = 'uppercase';

                if (cells[1]) {
                    cells[1].style.borderLeft = '5px solid #f59e0b';
                    cells[1].style.backgroundColor = 'rgba(245, 158, 11, 0.2)';
                    cells[1].style.color = '#fef3c7';
                    cells[1].style.fontWeight = '600';
                }

                coloredCount++;
            }

            // 🟢 NO NEED
            else if (text === 'NO NEED') {
                console.log(`🟢 NO NEED trouvé - ligne ${rowIndex}`);

                row.style.backgroundColor = 'rgba(16, 185, 129, 0.08)';

                cell.style.backgroundColor = 'rgba(16, 185, 129, 0.4)';
                cell.style.color = '#ffffff';
                cell.style.fontWeight = '700';
                cell.style.border = '3px solid #10b981';
                cell.style.borderRadius = '8px';
                cell.style.padding = '8px';
                cell.style.textTransform = 'uppercase';

                if (cells[1]) {
                    cells[1].style.borderLeft = '5px solid #10b981';
                }

                coloredCount++;
            }
        });
    });

    console.log(`✅ ${coloredCount} lignes colorées`);
}

// Exécution au chargement
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => {
        console.log("📄 DOM chargé, colorisation dans 500ms...");
        setTimeout(colorizeTableRows, 500);
    });
} else {
    console.log("📄 DOM déjà chargé, colorisation immédiate");
    setTimeout(colorizeTableRows, 500);
}

// Observer les changements
window.addEventListener('load', () => {
    console.log("🌐 Page complètement chargée");

    setTimeout(() => {
        const observer = new MutationObserver((mutations) => {
            console.log("🔄 Mutation détectée, recolorisation...");
            setTimeout(colorizeTableRows, 100);
        });

        const container = document.querySelector('.dash-table-container');
        if (container) {
            console.log("👀 Observer activé sur", container);
            observer.observe(container, {
                childList: true,
                subtree: true
            });
        } else {
            console.warn("⚠️ Container non trouvé pour observer");
        }
    }, 1000);
});

// Exposer la fonction globalement pour tests manuels
window.colorizeTableRows = colorizeTableRows;

console.log("✅ colorize.js prêt");

*/