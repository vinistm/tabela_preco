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
      var valorLiquido = jsonData[i][2].toString().replace("R$", "").trim();
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
});
