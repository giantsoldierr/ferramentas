document.getElementById('excel-upload').addEventListener('click', () => {
    document.getElementById('excel-input').click();
});

document.getElementById('docx-upload').addEventListener('click', () => {
    document.getElementById('docx-input').click();
});

document.getElementById('process-button').addEventListener('click', async () => {
    const excelFile = document.getElementById('excel-input').files[0];
    const docxFile = document.getElementById('docx-input').files[0];

    if (!excelFile || !docxFile) {
        alert('Por favor, envie os dois arquivos (Excel e DOCX)');
        return;
    }

    const formData = new FormData();
    formData.append('excel', excelFile);
    formData.append('docx', docxFile);

    // Leitura do arquivo Excel
    const reader = new FileReader();
    reader.onload = async function(e) {
        const data = e.target.result;
        const workbook = XLSX.read(data, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(sheet);

        // Leitura do modelo DOCX
        const docxReader = new FileReader();
        docxReader.onload = function(e) {
            const zip = new PizZip(e.target.result);
            const doc = new window.Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });

            const zipFile = new JSZip();

            jsonData.forEach((row) => {
                doc.setData(row);
                try {
                    doc.render();
                } catch (error) {
                    console.error("Erro ao renderizar documento:", error);
                    return;
                }

                const outputDoc = doc.getZip().generate({ type: 'arraybuffer' });
                zipFile.file(`${row['Nome']}.docx`, outputDoc);
            });

            // Gerar arquivo ZIP
            zipFile.generateAsync({ type: 'blob' }).then(function(content) {
                const link = document.createElement('a');
                link.href = URL.createObjectURL(content);
                link.download = 'documentos_gerados.zip';
                link.click();
                document.getElementById('result-text').textContent = 'Download concluído!';
            });
        };

        docxReader.readAsArrayBuffer(docxFile);
    };

    reader.readAsArrayBuffer(excelFile);
});
