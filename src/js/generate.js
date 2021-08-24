function loadFile(url, callback) {
  PizZipUtils.getBinaryContent(url, callback);
}

function GenerateS(url) {
  loadFile(url, function (error, content) {
    if (error) {
      throw error;
    }

    function replaceErrors(key, value) {
      if (value instanceof Error) {
        return Object.getOwnPropertyNames(value).reduce(function (error, key) {
          error[key] = value[key];
          return error;
        }, {});
      }
      return value;
    }
    function errorHandler(error) {
      console.log(JSON.stringify({ error: error }, replaceErrors));

      if (error.properties && error.properties.errors instanceof Array) {
        const errorMessages = error.properties.errors
          .map(function (error) {
            return error.properties.explanation;
          })
          .join("\n");
        console.log("errorMessages", errorMessages);
      }
      throw error;
    }

    var zip = new PizZip(content);
    var doc;
    try {
      doc = new window.docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
      });
    } catch (error) {
      errorHandler(error);
    }

    doc.setData({
      tutorial: document.getElementById("tutorial").value,
      direction: document.getElementById("direction").value,
      date: document.getElementById("date").value.split('-').reverse().join('.'),
      authors: document.getElementById("authors").value,
    });
    try {
      doc.render();
    } catch (error) {
      errorHandler(error);
    }

    var out = doc.getZip().generate({
      type: "blob",
      mimeType:
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });
    saveAs(out, "sluzhebnaya-zapiska.docx");
  });
}

function GenerateR(url) {
  loadFile(url, function (error, content) {
    if (error) {
      throw error;
    }

    function replaceErrors(key, value) {
      if (value instanceof Error) {
        return Object.getOwnPropertyNames(value).reduce(function (error, key) {
          error[key] = value[key];
          return error;
        }, {});
      }
      return value;
    }
    function errorHandler(error) {
      console.log(JSON.stringify({ error: error }, replaceErrors));

      if (error.properties && error.properties.errors instanceof Array) {
        const errorMessages = error.properties.errors
          .map(function (error) {
            return error.properties.explanation;
          })
          .join("\n");
        console.log("errorMessages", errorMessages);
      }
      throw error;
    }

    var zip = new PizZip(content);
    var doc;
    try {
      doc = new window.docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
      });
    } catch (error) {
      errorHandler(error);
    }

    doc.setData({
      institute: document.getElementById("institute").value,
      department: document.getElementById("department").value,
      tutorial: document.getElementById("tutorial").value,
      authors: document.getElementById("authors").value,
      title: document.getElementById("title").value,
      year: document.getElementById("year").value,
      direction: document.getElementById("direction").value,
      discipline: document.getElementById("discipline").value,
      email: document.getElementById("email").value,
      phone: document.getElementById("phone").value,
    });
    try {
      doc.render();
    } catch (error) {
      errorHandler(error);
    }

    var out = doc.getZip().generate({
      type: "blob",
      mimeType:
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });
    saveAs(out, "zayavka-na-razmeshchenie.docx");
  });
}
