<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Extrator de Parcelamentos PDF para Excel</title>
  
  <!-- Bibliotecas externas (PDF.js e SheetJS/XLSX) -->
  <script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
  
  <style>
    * { box-sizing: border-box; }
    body {
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
      max-width: 850px;
      margin: 40px auto;
      padding: 0 20px;
      background-color: #f4f6f8;
      color: #333;
    }
    .card {
      background: #ffffff;
      padding: 32px;
      border-radius: 12px;
      box-shadow: 0 4px 20px rgba(0,0,0,0.08);
    }
    h2 { margin-top: 0; color: #111827; }
    p { color: #4b5563; font-size: 15px; }
    .input-group {
      margin: 24px 0;
      padding: 20px;
      border: 2px dashed #cbd5e1;
      border-radius: 8px;
      background: #f8fafc;
      text-align: center;
    }
    input[type="file"] { margin-top: 8px; }
    button {
      background-color: #2563eb;
      color: white;
      border: none;
      padding: 12px 24px;
      font-size: 16px;
      border-radius: 6px;
      cursor: pointer;
      font-weight: 600;
      transition: background 0.2s;
      width: 100%;
    }
    button:hover { background-color: #1d4ed8; }
    button:disabled { background-color: #94a3b8; cursor: not-allowed; }
    #status {
      margin-top: 20px;
      font-weight: 600;
      color: #1e293b;
      word-break: break-word;
    }
    .note { font-size: 13px; color: #64748b; margin-top: 20px; text-align: center; }
  </style>
</head>
<body>

  <div class="card">
    <h2>Extrator de Parcelamentos PDF para Excel</h2>
    <p>Selecione até 50 PDFs de parcelamento para organizar todos os débitos em uma única planilha Excel.</p>
    
    <div class="input-group">
      <label for="pdfFiles"><strong>Escolha os arquivos PDF:</strong></label><br><br>
      <input type="file" id="pdfFiles" accept="application/pdf" multiple />
    </div>

    <button id="processBtn" onclick="processPDFs()">Processar e Baixar Planilha</button>

    <div id="status"></div>
    <div class="note">
      🔒 Processamento seguro: Todos os dados são lidos direto no seu navegador. Nenhum PDF é enviado para servidores externos.
    </div>
  </div>

  <script>
    pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

    async function processPDFs() {
      const input = document.getElementById('pdfFiles');
      const status = document.getElementById('status');
      const btn = document.getElementById('processBtn');

      if (!input.files.length) {
        alert("Por favor, selecione ao menos 1 arquivo PDF.");
        return;
      }

      if (input.files.length > 50) {
        alert("Por favor, selecione no máximo 50 arquivos por vez.");
        return;
      }

      btn.disabled = true;
      const allRows = [];

      for (let i = 0; i < input.files.length; i++) {
        const file = input.files[i];
        status.innerText = `⏳ Lendo arquivo ${i + 1} de ${input.files.length}: ${file.name}...`;

        try {
          const arrayBuffer = await file.arrayBuffer();
          const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
          
          let fullTextLines = [];

          for (let p = 1; p <= pdf.numPages; p++) {
            const page = await pdf.getPage(p);
            const textContent = await page.getTextContent();
            
            const linesMap = {};
            textContent.items.forEach(item => {
              const y = Math.round(item.transform[5]);
              if (!linesMap[y]) linesMap[y] = [];
              linesMap[y].push(item);
            });

            const sortedYs = Object.keys(linesMap).sort((a, b) => b - a);
            sortedYs.forEach(y => {
              const lineItems = linesMap[y].sort((a, b) => a.transform[4] - b.transform[4]);
              const lineText = lineItems.map(item => item.str.trim()).filter(Boolean);
              if (lineText.length > 0) {
                fullTextLines.push(lineText);
              }
            });
          }

          let parcelamentoNum = "Não identificado";
          for (const line of fullTextLines) {
            const joined = line.join(" ");
            const match = joined.match(/\b\d{10,20}\b/);
            if (match) {
              parcelamentoNum = match[0];
              break;
            }
          }

          fullTextLines.forEach(line => {
            if (line.length >= 8 && /^\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2}$/.test(line[0])) {
              allRows.push({
                "Nº Parcelamento": parcelamentoNum,
                "CNPJ do débito": line[0] || "",
                "Referência": line[1] && !line[1].includes('/') ? line[1] : "",
                "Processo Administrativo": line.find(str => str.includes('Proc')) || "",
                "Receita": line.find(str => /^\d{4}-\d{2}$/.test(str)) || "",
                "Período de apuração": line.find(str => str.includes('/20')) || "",
                "Vencimento": line.find(str => /^\d{2}\/\d{2}\/\d{4}$/.test(str)) || "",
                "Saldo originário": line.find(str => str.startsWith('BRL')) || "",
                "Principal (BRL)": line[line.length - 4] || "",
                "Multa (BRL)": line[line.length - 3] || "",
                "Juros (BRL)": line[line.length - 2] || "",
                "Valor consolidado (BRL)": line[line.length - 1] || "",
                "Arquivo Origem": file.name
              });
            }
          });

        } catch (err) {
          console.error(`Erro ao processar ${file.name}:`, err);
        }
      }

      if (allRows.length === 0) {
        status.innerText = "❌ Nenhum dado de parcelamento/tributo foi localizado nos PDFs selecionados.";
        btn.disabled = false;
        return;
      }

      status.innerText = "📊 Gerando planilha Excel...";
      const worksheet = XLSX.utils.json_to_sheet(allRows);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "Tributos Parcelados");

      XLSX.writeFile(workbook, `Tributos_Parcelamentos_${new Date().toISOString().slice(0,10)}.xlsx`);
      
      status.innerText = `✅ Sucesso! Total de ${allRows.length} linhas extraídas para o Excel.`;
      btn.disabled = false;
    }
  </script>

</body>
</html>
