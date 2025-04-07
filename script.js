function guardarRanking() {
    const comercial = document.getElementById('comercialRanking').value;
    const fecha = document.getElementById('fechaRanking').value;
    const ventasTotal = parseFloat(document.getElementById('ventasTotalRanking').value);

    if (!comercial || !fecha || !ventasTotal) {
        alert("Por favor, complete todos los campos.");
        return;
    }

    const rankingTable = document.getElementById('rankingTable').getElementsByTagName('tbody')[0];
    const newRow = rankingTable.insertRow();

    newRow.insertCell(0).textContent = comercial;
    newRow.insertCell(1).textContent = fecha;
    newRow.insertCell(2).textContent = ventasTotal;

    document.getElementById('rankingForm').reset();

    ordenarRanking();
}

function ordenarRanking() {
    const rows = Array.from(document.getElementById('rankingTable').getElementsByTagName('tbody')[0].rows);
    rows.sort((a, b) => parseFloat(b.cells[2].textContent) - parseFloat(a.cells[2].textContent));
    const rankingTable = document.getElementById('rankingTable').getElementsByTagName('tbody')[0];
    rankingTable.innerHTML = "";
    rows.forEach(row => rankingTable.appendChild(row));
}

function exportarExcel() {
    let csvContent = "data:text/csv;charset=utf-8,";
    const rows = document.querySelectorAll("#rankingTable tr");

    rows.forEach(row => {
        const cols = Array.from(row.children).map(col => col.textContent);
        csvContent += cols.join(",") + "\n";
    });

    const encodedUri = encodeURI(csvContent);
    const link = document.createElement("a");
    link.setAttribute("href", encodedUri);
    link.setAttribute("download", "ranking_comercial.csv");
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}