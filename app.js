let express = require('express');
let bodypar = require('body-parser');
const {GoogleSpreadsheet} = require('google-spreadsheet');
const docId = 'DOCID';
const doc = new GoogleSpreadsheet(docId);
doc.useServiceAccountAuth({
    "private_key": "KEY",
    "client_email": "EMAIL"
});  

let app = express();

app.listen(88);

app.use(bodypar.urlencoded({ extended:true}));
app.use(bodypar.json());

console.log('waiting');

async function readSheet(sheetIndex){
    let data = undefined;
    let descricao = undefined;
    let valor = undefined;
    let categoria = undefined;
    let natureza = undefined;  
    let transacao = {};
    let recordSet = {"transacao" : {}};

    //obtem as informações da planilha
    await doc.loadInfo();    
    const sheet = doc.sheetsByIndex[sheetIndex];
    await sheet.loadCells('A1:E300');

    let rows = await sheet.getRows();

    rows.forEach(function(row){
        data = sheet.getCellByA1('A' + row.rowIndex);
        descricao = sheet.getCellByA1('B' + row.rowIndex);
        valor = sheet.getCellByA1('C' + row.rowIndex);
        categoria = sheet.getCellByA1('D' + row.rowIndex);
        natureza = sheet.getCellByA1('E' + row.rowIndex);
        transacao = {
            "data" : data.formattedValue,
            "descricao" : descricao.formattedValue,
            "valor" : valor.formattedValue,
            "categoria" : categoria.formattedValue,
            "natureza" : natureza.formattedValue
        };

        recordSet.transacao[row.rowIndex] = transacao;
    });

    return recordSet;
    
}

async function writeSheet(sheetIndex, dados){

    //obtem as informações da planilha
    await doc.loadInfo();    
    const sheet = doc.sheetsByIndex[sheetIndex];
    await sheet.loadCells('A1:E100');

    let rows = await sheet.getRows();
    
    sheet.addRow(dados);

    return true;
    
}

app.get('/api/transacao', async function(req, res){
    let transacoes = await readSheet(2);
    res.send(transacoes);

});

app.post('/api/transacao', function(req, res){
    let dados = req.body;
    res.send(writeSheet(2, dados));
});