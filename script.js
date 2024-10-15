document.getElementById('searchButton').addEventListener('click', function() {
    const fileInput = document.getElementById('fileInput');
    const searchTerm = document.getElementById('searchInput').value.trim();
    
    // Si no se selecciona un archivo o el campo de búsqueda está vacío
    if (!fileInput.files.length) {
        alert('Por favor, selecciona un archivo Excel.');
        return;
    }
    
    if (searchTerm === "") {
        alert('Por favor, ingresa un término de búsqueda.');
        return;
    }

    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        // Variable para almacenar dónde se encontró el dato
        let foundInSheet = null;

        // Recorremos las 3 hojas: GCABA, PDC, IVC
        ['GCABA', 'PDC', 'IVC'].forEach(sheetName => {
            const sheet = workbook.Sheets[sheetName];
            if (sheet) {
                const jsonData = XLSX.utils.sheet_to_json(sheet);
                
                // Buscar el término dentro de la hoja
                const found = jsonData.some(row => 
                    Object.values(row).some(val => String(val).toLowerCase().includes(searchTerm.toLowerCase()))
                );
                
                if (found) {
                    foundInSheet = sheetName;
                }
            }
        });

        // Mostrar resultado en el div 'result'
        const resultDiv = document.getElementById('result');
        if (foundInSheet) {
            resultDiv.innerHTML = `Dato encontrado en la hoja: <strong>${foundInSheet}</strong>`;
            resultDiv.style.color = 'green';
        } else {
            resultDiv.innerHTML = 'Dato no encontrado en ninguna hoja.';
            resultDiv.style.color = 'red';
        }
    };

    reader.readAsArrayBuffer(fileInput.files[0]);
});
