const express = require("express");
const cors = require("cors");
const fs = require("fs");

const app = express();

app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(cors());

app.post('/salvaCorpoDoEmail', (req, res) => {
    if ([null, undefined, ''].includes(req.body["corpoDoEmail"])) {
        res.status(404).send("Sem corpo do email.");
    }

    fs.writeFile("corpoDoEmail.txt", req.body["corpoDoEmail"], (err) => {
        if (err) res.status(404).send("Erro ao salvar.");
    });

    res.send("Deu tudo certo.");
});

app.listen(8000, function () {
  console.log("Example app listening on port 8000!");
});