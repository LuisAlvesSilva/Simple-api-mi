async function initExcel() {
      const status = document.getElementById('statusMsg');
      try {
        const response = await fetch('Pasta1.xlsx');
        if (!response.ok) throw new Error(`Não foi possível carregar o arquivo Pasta1.xlsx (HTTP ${response.status})`);

        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });

        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const rows = XLSX.utils.sheet_to_json(worksheet);

        database = {};

        rows.forEach(row => {
          const marca = (row.Marca || row.MARCA || "").trim();
          
          // Ignora marcas que comecem com número (ex: "01", "123", etc.)
          if (/^\d/.test(marca)) return;

          const modelo = (row.Modelo || row.MODELO || "").trim();

          const edicao = (row.Edicao || row.Edição || row.EDICAO || "GERAL").trim();
          
          const nomeArquivo = (row.Name || row.PN || row.Arquivo || "").trim();
          
          let url = (row.Url || row.URL || row["Folder Path"] || "").trim();
          if (url && nomeArquivo && !url.endsWith('/')) {
            url = url + '/' + nomeArquivo;
          } else if (url && nomeArquivo) {
            url = url + nomeArquivo;
          }

          if (!marca || !modelo || !nomeArquivo) return;

          let pn = nomeArquivo.replace(/\.[^/.]+$/, ""); 
          const pnMatch = nomeArquivo.match(/\b\d{7,8}\b/);
          if (pnMatch) pn = `PN ${pnMatch[0]} (${pn})`;

          if (!database[marca]) database[marca] = {};
          if (!database[marca][modelo]) database[marca][modelo] = {};
          if (!database[marca][modelo][edicao]) database[marca][modelo][edicao] = [];

          database[marca][modelo][edicao].push({ pn, url });
        });

        // Ordena os itens do PN de forma crescente alfabeticamente
        Object.keys(database).forEach(marca => {
          Object.keys(database[marca]).forEach(modelo => {
            Object.keys(database[marca][modelo]).forEach(edicao => {
              database[marca][modelo][edicao].sort((a, b) => a.pn.localeCompare(b.pn, 'pt-BR'));
            });
          });
        });

        status.textContent = "Base carregada com sucesso do Excel!";
        listBrand();

      } catch (err) {
        console.error(err);
        status.innerHTML = `Erro: ${err.message}. Certifique-se de rodar via Live Server.`;
      }
    }