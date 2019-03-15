//<------------------------------------------All Required Imports--------------------------------------->
const readXlsxFile = require("read-excel-file");
const input = document.getElementById("input");
const fs = require("fs");
const pdf = require("html-pdf");

//<------------------------------------------Global Variables--------------------------------------->
var allExcelRows = [];
let options = { format: "Letter" };
let destinationPath = ".";
const nameLabels = {
  Name: "NAME",
  Address: "",
  City: "",
  InvoiceNumber: "",
  Content: "",
  Quantity: "",
  SpecialIntructions: "",
  Mobile1: "",
  COD: ""
};

//<------------------------------------------Event linsters--------------------------------------->

/**
 * Excel file change event listener
 */
document.getElementById("excelInput").addEventListener("change", () => {
  readXlsxFile(document.getElementById("excelInput").files[0], {
    sheet: 1
  }).then(rows => {
    console.log(rows);
    const colIndexMapping = getColumnsIndexMapping(rows[0]);
    console.log(colIndexMapping);

     pdf.create(generateAllInvoices(rows), options).toFile(document.getElementById("destinationPath").files[0].path+'/businesscard.pdf', function(err, res) {
        if (err) return console.log(err);
        console.log(res);
      });
    // `rows` is an array of rows
    // each row being an array of cells.
  });

  /**
   * Destination folder change event listener
   */
  document.getElementById("destinationPath").addEventListener("change", () => {
    destinationPath = document.getElementById("destinationPath").files[0].path;
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
  function getDivider(){
    return `
      <div style="width:100%;border-bottom:1px dotted black"></div>
    `;
  }

  /**
   * generate all invoices
   */
  function generateAllInvoices(rows){
    let html="";
    rows.forEach((v,i)=>{
      const pageBreakStyle=i==0?'':"page-break-before:always;";
      html+=`<div style="${pageBreakStyle}">${generateSingleInvoice()}${getDivider()}${generateSingleInvoice()}</div>`
    });
    return html;
  }


  /**
   * generate single invoice
   */
  function generateSingleInvoice(row, colIndexMapping) {
    return `
    <div class="invDiv" style="padding:20px;margin-top:50px;">
    <div class="invHeader" style="padding-bottom: 10px; border-bottom: 1px solid black;">
        <table style="font-weight:bold;width:100% ">
            <tr>
                <td style="width:70%;font-size:21px;color:#186ba0">
                    FID General Trading FZE
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
                    Invoice No: 23124123
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
                    <br /> CheckDeals
                    <br /> Ajman,
                    <br />
                    <br />Ajman
                    <br />3412524123551234
                </td>
                <td style="width:50%;">
                    <table style="width:100%;font-size:13px;">
                        <tr>
                            <td style="width:50%;padding-left:5%;font-weight:bold;">
                                Date:
                            </td>
                            <td style="width:50%;">
                                12-Mar-2334
                            </td>
                        </tr>
                        <tr>
                            <td style="width:50%;padding-left:5%;font-weight:bold;">
                                Total Pieces:
                            </td>
                            <td style="width:50%;">
                                32
                            </td>
                        </tr>
                        <tr>
                            <td style="width:50%;padding-left:5%;font-weight:bold;">
                                COD Amount:
                            </td>
                            <td style="width:50%;">
                                323
                            </td>
                        </tr>
                        <tr>
                            <td style="width:50%;padding-left:5%;font-weight:bold;">
                                Special Instructions:
                            </td>
                            <td style="width:50%;">
                                OK os fao asid aosdfo asd
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
                    <br /> Name
                    <br />
                    <br /> al ameriya lady couture near al jimi mall al khaleeje hospital besides al raqi meeical clinic.alain
                    abudhabi,
                    <br />
                    <br />Ajman
                    <br />3412524123551234
                </td>
                <td style="width:50%;padding:0% 3%;border-top:1px solid black;">
                    <b>Description of Goods:</b>
                    <br />
                    <br /> fajsdf la skjflafjf;af fla jflk sf fkklj fa dkf alsfa sfkj alkf ljask fa iodfao
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
