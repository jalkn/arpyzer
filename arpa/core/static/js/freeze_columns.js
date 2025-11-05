document.addEventListener('DOMContentLoaded', function() {
    const table = document.querySelector('.table');
    if (!table) return;

    const freezeButtons = document.querySelectorAll('.freeze-column-btn');
    let frozenColumns = JSON.parse(localStorage.getItem('frozenColumns')) || [];

    function applyFrozenColumns() {
        // Clear any existing frozen classes and inline styles
        document.querySelectorAll('.table-frozen-column').forEach(el => {
            el.classList.remove('table-frozen-column');
            el.style.left = ''; // Clear inline style
        });
        document.querySelectorAll('.freeze-column-btn').forEach(btn => {
            btn.classList.remove('active');
        });

        let currentLeft = 0;
        frozenColumns.forEach(colIndex => {
            const cellsInColumn = table.querySelectorAll(`td:nth-child(${colIndex + 1}), th:nth-child(${colIndex + 1})`);
            cellsInColumn.forEach(cell => {
                cell.classList.add('table-frozen-column');
                cell.style.left = `${currentLeft}px`;
            });

            // Mark the corresponding freeze button as active
            const button = document.querySelector(`.freeze-column-btn[data-column-index="${colIndex}"]`);
            if (button) {
                button.classList.add('active');
            }

            // Calculate the width of the frozen column to offset the next one
            // This is a simplified approach, in a real complex table with variable widths,
            // you might need a more robust calculation or a library.
            const headerCell = table.querySelector(`th:nth-child(${colIndex + 1})`);
            if (headerCell) {
                currentLeft += headerCell.offsetWidth;
            }
        });
    }

    freezeButtons.forEach(button => {
        button.addEventListener('click', function() {
            const columnIndex = parseInt(this.dataset.columnIndex);
            const indexInFrozen = frozenColumns.indexOf(columnIndex);

            if (indexInFrozen > -1) {
                // Column is already frozen, unfreeze it
                frozenColumns.splice(indexInFrozen, 1);
            } else {
                // Column is not frozen, freeze it
                frozenColumns.push(columnIndex);
                frozenColumns.sort((a, b) => a - b); // Keep columns ordered by index
            }

            localStorage.setItem('frozenColumns', JSON.stringify(frozenColumns));
            applyFrozenColumns();
        });
    });

    // Apply frozen columns on initial load
    applyFrozenColumns();

    // Re-apply frozen columns on window resize to adjust 'left' positions
    window.addEventListener('resize', applyFrozenColumns);
});
