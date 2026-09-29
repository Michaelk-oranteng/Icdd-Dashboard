// responsive-tables.js
(function() {
    'use strict';
    
    class ResponsiveTables {
        constructor() {
            this.tables = document.querySelectorAll(
                '.activity-table, .users-table, .checklist-table, .deductions-table, .report-table'
            );
            this.init();
        }
        
        init() {
            this.tables.forEach(table => this.setupTable(table));
            
            // Re-setup on resize
            let resizeTimer;
            window.addEventListener('resize', () => {
                clearTimeout(resizeTimer);
                resizeTimer = setTimeout(() => {
                    this.tables.forEach(table => this.setupTable(table));
                }, 150);
            });
        }
        
        setupTable(table) {
            const headers = table.querySelectorAll('thead th');
            if (!headers.length) return;
            
            const headerLabels = Array.from(headers).map(th => 
                th.textContent.trim() || th.dataset.label || ''
            );
            
            const rows = table.querySelectorAll('tbody tr');
            
            rows.forEach(row => {
                // Skip empty rows
                if (row.querySelector('.empty-row, .users-table-empty')) return;
                
                const cells = row.querySelectorAll('td');
                cells.forEach((cell, index) => {
                    if (headerLabels[index]) {
                        cell.setAttribute('data-label', headerLabels[index]);
                    }
                });
            });
        }
    }
    
    // Initialize
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', () => new ResponsiveTables());
    } else {
        new ResponsiveTables();
    }
})();