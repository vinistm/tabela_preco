function handleFile(e) {
  var codigoTabPreco = document.getElementById("codigoTabPreco").value;
  if (!codigoTabPreco) {
    alert("Por favor, selecione uma tabela de preço.");
    return;
  }

  var file = e.target.files[0];
  var reader = new FileReader();

  reader.onload = function (e) {
    var data = new Uint8Array(e.target.result);
    var workbook = XLSX.read(data, { type: "array" });
    var sheet = workbook.Sheets[workbook.SheetNames[0]];
    var jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    var valorOption = document.getElementById("valorOption").value;
    var outputText = "";
    for (var i = 1; i < jsonData.length; i++) {
      var produto = jsonData[i][0];
      var cor = jsonData[i][1];
      var valorLiquido = Number(
        jsonData[i][2].toString().replace("R$", "").trim()
      ).toFixed(2);
      var updateQuery = "";
      if (valorOption === "valor_liquido") {
        updateQuery =
          "update produtos_precos set preco_liquido1='" +
          valorLiquido +
          "' where produto='" +
          produto +
          "' and codigo_tab_preco='" +
          codigoTabPreco +
          "';";
      } else if (valorOption === "preco_valor_liquido") {
        updateQuery =
          "update produtos_precos set preco1='" +
          valorLiquido +
          "', preco_liquido1='" +
          valorLiquido +
          "' where produto='" +
          produto +
          "' and codigo_tab_preco='" +
          codigoTabPreco +
          "';";
      } else if (valorOption === "preco_valor_cor") {
        updateQuery =
          "update produtos_preco_cor set preco1='" +
          valorLiquido +
          "', preco_liquido1='" +
          valorLiquido +
          "' where produto='" +
          produto +
          "' and cor_produto='" +
          cor +
          "' and codigo_tab_preco='" +
          codigoTabPreco +
          "';";
      } else if (valorOption === "preco_valor_cor_insert") {
        updateQuery =
          "INSERT INTO PRODUTOS_PRECO_COR(CODIGO_TAB_PRECO,PRODUTO,COR_PRODUTO,PRECO1,PRECO_LIQUIDO1) VALUES('" +
          codigoTabPreco +
          "', '" +
          produto +
          "','" +
          cor +
          "','" +
          valorLiquido +
          "','" +
          valorLiquido +
          "')";
      }
      outputText += updateQuery + "\n";
    }

    var outputTextarea = document.getElementById("output");
    outputTextarea.value = outputText;
  };

  reader.readAsArrayBuffer(file);
}

document
  .getElementById("inputFile")
  .addEventListener("change", handleFile, false);

document
  .getElementById("generateButton")
  .addEventListener("click", function () {
    var outputTextarea = document.getElementById("output");
    var modal = new bootstrap.Modal(document.getElementById("myModal"));
    modal.show();
  });

document.getElementById("copyButton").addEventListener("click", function () {
  var outputTextarea = document.getElementById("output");
  outputTextarea.select();
  document.execCommand("copy");
  alert("Texto copiado com sucesso!");
});

document.getElementById("reloadButton").addEventListener("click", function () {
  location.reload();
  alert("Selecione nova tabela e suba um novo arquivo!");
});
