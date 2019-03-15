//<------------------------------------------All Required Imports--------------------------------------->
const readXlsxFile = require("read-excel-file");
const input = document.getElementById("input");
const fs = require("fs");
const pdf = require("html-pdf");

//<------------------------------------------Global Variables--------------------------------------->
let options = { format: "Letter" };
let destinationPath = ".";
let allExcelRows = [];
let colIndexMapping = {};
const nameLabels = {
  Name: "NAME",
  Address: "ADDRESS",
  City: "CITY",
  InvoiceNumber: "INVOICE NO",
  Content: "CONTENT",
  Quantity: "QTY",
  SpecialIntructions: "SPECIAL INSTRUCTION",
  Mobile1: "MOBILE 1",
  COD: "COD"
};

//<------------------------------------------Event linsters--------------------------------------->

/**
 * Excel file change event listener
 */
document.getElementById("excelInput").addEventListener("change", () => {
  readXlsxFile(document.getElementById("excelInput").files[0], {
    sheet: 1
  }).then(rows => {
    try {
      console.log(rows);
      allExcelRows = rows;
      colIndexMapping = getColumnsIndexMapping(rows[0]);
      console.log(colIndexMapping);
      allExcelRows.splice(0, 1);
    } catch (ex) {
      alert("Error occured while reading excel file.\nError: " + ex.message);
    }

     
  });

  /**
   * Destination folder change event listener
   */
  document.getElementById("destinationPath").addEventListener("change", () => {
    destinationPath = document.getElementById("destinationPath").files[0].path;
  });

  /**
   * Generate pdf button event listener
   */
  document.getElementById("generateButton").addEventListener("click", () => {
    document.getElementById("generateLoader").style.display = "block";
    setTimeout(() => {
      try {
        pdf
          .create(generateAllInvoices(allExcelRows), options)
          .toFile(
            document.getElementById("destinationPath").files[0].path +
              "/" +
              document.getElementById("fileName").value +
              ".pdf",
            function(err, res) {
              if (err) return console.log(err);
              console.log(res);
              document.getElementsByClassName("loader")[0].style.display =
                "none";
                alert("Pdf generated successfully. Please check the desnitation folder.")
            }
          );
      } catch (ex) {
        alert("Error occured while generating pdf.\nError: " + ex.message);
      }
    }, `500`);
  });

  //<------------------------------------------Gloabal Functions--------------------------------------->

  /**
   * validateAllInputs
   */
  function validateInputs() {
    if (
      !document.getElementById("shipperName") ||
      !document.getElementById("shipperAddress") ||
      !document.getElementById("shipperContact") ||
      !document.getElementById("invoiceDate")
    )
      return false;
    return true;
  }

  /**
   * PDF html generator
   */
  function generatePDF() {
    let pdfHTML = "";
  }

  /**
   * gets divider for two receipts
   */
  function getDivider() {
    return `
      <div style="width:100%;border-bottom:1px dotted black;margin-top:20px;"></div>
    `;
  }

  /**
   * generate all invoices
   */
  function generateAllInvoices(rows) {
    let html = "";
    rows.forEach((v, i) => {
      const pageBreakStyle = i == 0 ? "" : "page-break-before:always;";
      html += `<div style="${pageBreakStyle}">${generateSingleInvoice(
        rows[i]
      )}${getDivider()}${generateSingleInvoice(rows[i])}</div>`;
    });
    return html;
  }

  /**
   * generate single invoice
   */
  function generateSingleInvoice(row) {
    return `
    <div class="invDiv" style="padding:20px;margin-top:50px;">
    <div class="invHeader" style="padding-bottom: 10px; border-bottom: 1px solid black;">
        <table style="font-weight:bold;width:100% ">
            <tr>
                <td style="width:70%;font-size:21px;color:#186ba0">
                    FID General Trading FZC
                </td>
                <td style="width:30%;text-align:right;">
                    Tax Invoice
                </td>
            </tr>
            <tr>
                <td style="width:70%;font-size:19px;">
                    TRN No: 100460738600003
                </td>
                <td style="width:30%;text-align:right;font-weight:normal;">
                    Invoice No: ${row[colIndexMapping[nameLabels["InvoiceNumber"]]]}
                </td>
            </tr>
        </table>
    </div>
    <div class="invBody">
        <table style="border-bottom:1px solid black;font-size: 13px;width: 100%;">
            <tr>
                <td style="width:50%;padding-bottom:15px;border-right:1px solid black;padding-top:15px;">
                    <b>
                        <u>SHIPPER</u>
                    </b>
                    <br /> ${document.getElementById("shipperName").value}
                    <br /> ${document.getElementById("shipperAddress").value},
                    <br />
                    <br />${document.getElementById("shipperAddress").value}
                    <br />${document.getElementById("shipperContact").value}
                </td>
                <td style="width:50%;">
                    <table style="width:100%;font-size:13px;">
                        <tr>
                            <td style="width:50%;padding-left:5%;font-weight:bold;">
                                Date:
                            </td>
                            <td style="width:50%;">
                                ${document.getElementById("invoiceDate").value}
                            </td>
                        </tr>
                        <tr>
                            <td style="width:50%;padding-left:5%;font-weight:bold;">
                                Total Pieces:
                            </td>
                            <td style="width:50%;">
                            ${row[colIndexMapping[nameLabels["Quantity"]]]}
                            </td>
                        </tr>
                        <tr>
                            <td style="width:50%;padding-left:5%;font-weight:bold;">
                                COD Amount:
                            </td>
                            <td style="width:50%;">
                            ${row[colIndexMapping[nameLabels["COD"]]]}
                            </td>
                        </tr>
                        <tr>
                            <td style="width:50%;padding-left:5%;font-weight:bold;">
                                Special Instructions:
                            </td>
                            <td style="width:50%;">
                            ${row[colIndexMapping[nameLabels["SpecialIntructions"]]]}
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td style="width:50%;border-top:1px solid black;border-right:1px solid black;padding-top:15px;">
                    <b>
                        <u>CONSIGNEE</u>
                    </b>
                    <br /> ${row[colIndexMapping[nameLabels["Name"]]]}
                    <br />
                    <br /> ${row[colIndexMapping[nameLabels["Address"]]]}
                    <br />
                    <br />${row[colIndexMapping[nameLabels["City"]]]}
                    <br />${row[colIndexMapping[nameLabels["Mobile1"]]]}
                </td>
                <td style="width:50%;padding:0% 3%;border-top:1px solid black;">
                    <b>Description of Goods:</b>
                    <br />
                    <br /> ${row[colIndexMapping[nameLabels["Content"]]]}
                    <div style="margin-top: 25px;
                    border-top: 1px solid black;
                    padding-top: 10px;">Received by: _________________________</div>
                </td>
            </tr>
        </table>
    </div>
</div>
    `;
  }

  /**
   * Get columns to index mapping
   */
  function getColumnsIndexMapping(firstRow) {
    let mapping = {};
    firstRow.forEach((v, i) => {
      mapping[v] = i;
    });
    return mapping;
  }
});
