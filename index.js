const ExcelJS = require('exceljs');
const express = require('express');
const app = express();

const meses = ["JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"];

// --- DESIGN SYSTEM (Tema: Luxury Black & Gold) ---
const PALETA = {
    fundoPreto: 'FF000000',       // Preto puro
    douradoOuro: 'FFD4AF37',      // Dourado metálico clássico
    douradoSuave: 'FFFFF2CC',     // Um tom creme/dourado bem clarinho para alternar linhas (opcional)
    textoBranco: 'FFFFFFFF',      // Branco
    textoPreto: 'FF000000',       // Preto
    bordaDourada: 'FFB8860B'      // Dark Goldenrod (um dourado mais escuro para as bordas ficarem visíveis)
};

const estilos = {
    fillSolido: (corArgb) => ({ type: 'pattern', pattern: 'solid', fgColor: { argb: corArgb } }),
    bordaFina: { style: 'thin', color: { argb: PALETA.bordaDourada } },
    alinhamentoCentro: { vertical: 'middle', horizontal: 'center', wrapText: true },
    alinhamentoEsquerda: { vertical: 'middle', horizontal: 'left' }
};
estilos.bordaCompleta = { top: estilos.bordaFina, left: estilos.bordaFina, bottom: estilos.bordaFina, right: estilos.bordaFina };


async function gerarPlanilha(empresa) {
    const workbook = new ExcelJS.Workbook();
    
    // showGridLines: false faz as células vazias em volta ficarem brancas (sem linhas)
    const sheet = workbook.addWorksheet('Faturamento 2026', { 
        views: [{ showGridLines: false }],
        pageSetup: { horizontalCentered: true, verticalCentered: false } // Centraliza na impressão
    });

    // --- SEÇÃO 1: CABEÇALHO (PRETO COM TEXTO DOURADO) ---
    sheet.getCell('A1').value = "ANO 2026";
    sheet.getCell('A2').value = `FIRMA: ${empresa.nome.toUpperCase()}`;
    sheet.getCell('A3').value = `TELEFONE: ${empresa.telefone || ''}`;
    sheet.getCell('A4').value = `CNPJ: ${empresa.cnpj}`;
    sheet.getCell('A5').value = `INSCRIÇÃO ESTADUAL: ${empresa.inscricao}`;
    sheet.getCell('A6').value = `SENHA EMISSÃO NFS-E: ${empresa.senha_nfse || ''}`;

    for (let r = 1; r <= 6; r++) {
        sheet.mergeCells(`A${r}:G${r}`);
        const cell = sheet.getCell(`A${r}`);
        
        cell.fill = estilos.fillSolido(PALETA.fundoPreto);
        cell.font = { 
            name: 'Segoe UI', 
            size: (r === 1 ? 16 : 11), 
            bold: true, 
            color: { argb: PALETA.douradoOuro } // Texto Dourado no fundo Preto
        };
        cell.alignment = { ...estilos.alinhamentoEsquerda, indent: 1 };
        cell.border = { 
            left: estilos.bordaFina, 
            right: estilos.bordaFina,
            top: (r === 1 ? estilos.bordaFina : undefined), // Borda só no topo da primeira
            bottom: (r === 6 ? estilos.bordaFina : undefined) // Borda só no fundo da última
        };
    }

    // --- SEÇÃO 2: TÍTULOS (DOURADO COM TEXTO PRETO) ---
    sheet.mergeCells('A8:A9'); sheet.getCell('A8').value = "MÊS";
    sheet.mergeCells('B8:C8'); sheet.getCell('B8').value = "FATURAMENTO";
    sheet.mergeCells('D8:E8'); sheet.getCell('D8').value = "VENDAS DE SERVIÇOS\n(c/ e s/ Retenção)";
    sheet.mergeCells('F8:F9'); sheet.getCell('F8').value = "%";
    sheet.mergeCells('G8:G9'); sheet.getCell('G8').value = "VALOR DO IMPOSTO";

    sheet.getCell('B9').value = "Mensal";
    sheet.getCell('C9').value = "Acumulado";
    sheet.getCell('D9').value = "C/ Retenção";
    sheet.getCell('E9').value = "S/ Retenção";

    sheet.getRow(8).height = 35;
    [8, 9].forEach(r => {
        sheet.getRow(r).eachCell({ includeEmpty: false }, (cell, colNumber) => {
            if (colNumber <= 7) {
                cell.fill = estilos.fillSolido(PALETA.douradoOuro); // Fundo Dourado
                cell.font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: PALETA.textoPreto } }; // Texto Preto
                cell.alignment = estilos.alinhamentoCentro;
                cell.border = estilos.bordaCompleta;
            }
        });
    });

    // --- SEÇÃO 3: CORPO DA TABELA ---
    meses.forEach((mes, index) => {
        const rowNum = 10 + index;
        const row = sheet.getRow(rowNum);
        row.height = 20;

        // Mês (Coluna A) - Destaque em Dourado também
        const cellMes = row.getCell(1);
        cellMes.value = mes;
        cellMes.font = { name: 'Segoe UI', bold: true, color: { argb: PALETA.textoPreto } };
        cellMes.fill = estilos.fillSolido(PALETA.douradoOuro);
        
        // Fórmulas
        row.getCell(3).value = { formula: index === 0 ? `IF(B${rowNum}="","",B${rowNum})` : `IF(B${rowNum}="","",B${rowNum}+C${rowNum - 1})` };
        row.getCell(7).value = { formula: `IF(F${rowNum}="","",(D${rowNum}+E${rowNum})*F${rowNum})` };

        // Formatação das Células
        for (let i = 1; i <= 7; i++) {
            const cell = row.getCell(i);
            cell.border = estilos.bordaCompleta; // Bordas Douradas
            
            // Alinhamento
            cell.alignment = estilos.alinhamentoCentro;

            // Formatos
            if ([2, 3, 4, 5, 7].includes(i)) cell.numFmt = '_-R$ * #,##0.00_-';
            if (i === 6) cell.numFmt = '0.00%';
        }
    });

    // --- AJUSTE DE LARGURA ---
    sheet.getColumn('A').width = 12;
    sheet.getColumn('B').width = 18;
    sheet.getColumn('C').width = 18;
    sheet.getColumn('D').width = 16;
    sheet.getColumn('E').width = 16;
    sheet.getColumn('F').width = 8;
    sheet.getColumn('G').width = 20;

    return await workbook.xlsx.writeBuffer();
}

// --- SERVIDOR WEB ---
app.use(express.static('public')); // Serve os arquivos da pasta 'public'
app.use(express.json());

app.get('/', (req, res) => {
    res.sendFile(__dirname + '/index.html');
});

app.post('/gerar', async (req, res) => {
    try {
        const buffer = await gerarPlanilha(req.body);
        const nomeLimpo = (req.body.nome || 'Empresa').replace(/[^a-zA-Z0-9]/g, '_').toUpperCase();
        
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename=${nomeLimpo}_FATURAMENTO_2026.xlsx`);
        res.send(buffer);
        console.log(`🖤💛 Planilha gerada para: ${req.body.nome}`);
    } catch (error) {
        console.error(error);
        res.status(500).send('Erro ao gerar planilha');
    }
});

app.listen(3000, () => {
    console.log('🖤💛 Sistema Luxury rodando em: http://localhost:3000');
});