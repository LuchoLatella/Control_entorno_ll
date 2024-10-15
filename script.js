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

        let foundInSheet = null;
        let foundRow = null;  // Aquí almacenaremos la fila completa donde se encontró el dato

        // Recorremos las 3 hojas: GCABA, PDC, IVC
        ['GCABA', 'PDC', 'IVC'].forEach(sheetName => {
            const sheet = workbook.Sheets[sheetName];
            if (sheet) {
                const jsonData = XLSX.utils.sheet_to_json(sheet);
                
                // Buscar el término dentro de la hoja
                jsonData.forEach(row => {
                    if (Object.values(row).some(val => String(val).toLowerCase().includes(searchTerm.toLowerCase()))) {
                        foundInSheet = sheetName;
                        foundRow = row;  // Guardamos la fila encontrada
                    }
                });
            }
        });

        const resultDiv = document.getElementById('result');
        resultDiv.innerHTML = ''; // Limpiar resultados anteriores

        if (foundInSheet && foundRow) {
            if (foundInSheet === 'GCABA') {
                resultDiv.innerHTML = `
                    Corresponde al entorno: <strong>${foundInSheet}</strong><br>
                    CUIL: ${foundRow['CUIL']}<br>
                    CARGO: ${foundRow['CARGO']}<br>
                    AYN: ${foundRow['AYN']}<br>
                    COD_REP: ${foundRow['COD_REP']}<br>
                    DESC_REP: ${foundRow['DESC_REP']}<br>
                    MINISTERIO: ${foundRow['MINISTERIO']}<br>
                    CAR_SIT_REV: ${foundRow['CAR_SIT_REV']}
                `;
                resultDiv.style.color = 'green';
            } else if (foundInSheet === 'PDC') {
                resultDiv.innerHTML = `
                    Corresponde al entorno: <strong>${foundInSheet}</strong><br>
                    CUIL: ${foundRow['CUIL']}<br>
                    CARGO: ${foundRow['CARGO']}<br>
                    AYN: ${foundRow['AYN']}<br>
                    COD_REP: ${foundRow['COD_REP']}<br>
                    DESC_REP: ${foundRow['DESC_REP']}<br>
                    MINISTERIO: ${foundRow['MINISTERIO']}
                `;
                resultDiv.style.color = 'green';
            } else if (foundInSheet === 'IVC') {
                resultDiv.innerHTML = `
                    Corresponde al entorno: <strong>${foundInSheet}</strong><br>
                    CUIL: ${foundRow['CUIL']}<br>
                    CARGO: ${foundRow['CARGO']}<br>
                    AYN: ${foundRow['AYN']}<br>
                    COD_REP: ${foundRow['COD_REP']}<br>
                    DESC_REP: ${foundRow['DESC_REP']}<br>
                    MINISTERIO: ${foundRow['MINISTERIO']}
                `;
                resultDiv.style.color = 'green';
            }
        } else {
            resultDiv.innerHTML = 'Lo siento, dato no encontrado.';
            resultDiv.style.color = 'red';
        }
    };

    reader.readAsArrayBuffer(fileInput.files[0]);
});
