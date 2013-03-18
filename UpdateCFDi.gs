// Copyright (c) 2013, Lae,
//                     Enrique Gámez <egamezf@gmail.com>
//
// All rights reserved.
//
// Redistribution and use in source and binary forms, with or without
// modification, are permitted provided that the following conditions are met:
//
// - Redistributions of source code must retain the above copyright notice,
//   this list of conditions and the following disclaimer.
// - Redistributions in binary form must reproduce the above copyright notice,
//   this list of conditions and the following disclaimer in the documentation
//   and/or other materials provided with the distribution.
//
// THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
// AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
// IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
// ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE
// LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
// CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
// SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
// INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
// CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
// ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
// POSSIBILITY OF SUCH DAMAGE.
//
// The main pourpouse of this script is to organize the information,
// sent as file attachments on an emai, corresponding to a electronic
// invoices in Mexico (CFD or CFDi).
//
// With organize I meant,
//
//     - Save the CFD instance, and possibly its printed version
//       on Google Drive,
//     - Archive the email under some defined labels, and
//     - Store some brief information of the CFD on a Google Spreadsheet.
//
// The default behavior of the script is to store all the file attachments
// under the folder "Compras" (through the variable 'root_folder'), archive
// all the messages under the label "Gastos" (variable 'default_label') on
// the mailbox, and to create the spreadsheets with the prefix name
// "Reporte-Compras" (variable 'root_name'). The user can modify those
// names through the values of the variables mentioned.
//
// If the user wants to archive the messages under some different mailbox
// labels (one per message) can define those labels using User Properties.
// The rule to do so, it is by creating a property with the R.F.C. as the
// key, and the label name as its value, i.e. if you create a property
// with:
//
//        Key: AAAA010101AA
//      Value: Oficina
//
// the script will archive all the CFDs which emisor has R.F.C. 'AAAA010101AAA'
// under the mailbox label 'Oficina'.
//
// All the CFD instances (and their printend versions, if any) will be
// saved on Google Drive, under the folder name, given as the value of
// the variable 'root_folder', and the structure will be the following:
//
//      [root_folder]/
//            [emision-year-number]/
//                    [emision-month-number]
//
// The 'emision-year-number' and 'emision-month-number' will be inferred
// from the CFD itself, and will be created (if necesary) on the fly.
//
// The spreadsheet, that will also be created, will contain the following
// information (one row, per CDF succesfully parsed):
//
//            Information content:                         Column Label:
//       Fecha de emision del comprobante                [Fecha emision]
//       Version del comprobante                         [Version]
//       Serie del comprobante                           [Serie]
//       Folio del comprobante                           [Folio]
//       Folio fiscal del CFDi                           [Folio Fiscal]
//       Nombre de el proveedor                          [Proveedor]
//       R.F.C. del proveedor                            [R.F.C.]
//       Tipo de comprobante (Factura, Nota de Credito)  [Tipo]
//       Monto total de los descuentos del comprobante   [Descuentos]
//       Subtotal                                        [Subtotal]
//       I.V.A.                                          [I.V.A.]
//       Total                                           [Total]
//       Fecha de recepcion del comprobante              [Recepcion]
//       Fecha de recepcion de las mercancias            [Recepcion mercancia]
//       Fecha de pago del comprobante                   [Fecha pago]
//       Monto del pago                                  [Monto de pago]
//       Observaciones                                   [Observaciones]
//       Estado del comprobante (Vigente o Cancelado)    [Estado]
//       Url del comprobante (printed version if any)    [Url]
//       Col1 (columna para usos variables)              [Col1]
//
// This spreadsheet will be created on the fly, one per month. It is
// worth to mention that User Properties are used to save the the
// Spreadsheet keys of all the sheets created. No user interaction
// is required; the User Properties created for this pourpose have
// the signature:
//
//          [root_name]-[year]-[month]-key
//
// Bugs and comments to: egamezf@gmail.com
//

var root_folder   = "Compras"; // Root folder name under we will save the CFDs
var root_name     = "Reporte-Compras"; // Base name for the sheets created.
var default_label = "Gastos"; // Default Gmail label to archive the messages

function update_cfdi() {

  // Get the inbox thread
  var threads = GmailApp.getInboxThreads();

  // Loop over all the messages in the inbox thread.
  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();

    // Inspect every message in the collection if they has any attachment.
    for (var j = 0; j < messages.length; j++) {

      // Just parse those messages with attachments
      if ( messages[j].getAttachments().length > 0 ) {
        // Work out this message. We need to send
        // the attachments to Google Drive onto the
        // appropiate folder, and archive it with
        // the appropiate label on the mailbox.

        // Parse the attachments, maybe some of them are CFD/CFDis
        var cfdis = parse_cfdi(messages[j]);

        for (var k = 0; k < cfdis.length; k++) {

          if ( cfdis[k].comprobante === null ) {
            // Make as if you haven't touch the message.
            messages[j].markUnread();
            continue;
          }

          // Copy the current attachment to the given folder
          var ano = cfdis[k].fecha.getFullYear().toString();
          var mes = cfdis[k].fecha.getMonth() + 1;
          var dia = cfdis[k].fecha.getDate().toString();

          // Check that the folders exist, otherwise creat it.
          var folder = load_folder(messages[j], root_folder, ano, mes.toString());

          var cfdi_tipo = undefined;
          if ( cfdis[k].tipo == 1 ) cfdi_tipo = "Factura";
          else                      cfdi_tipo = "Nota de Crédito";

          var recepcion = messages[j].getDate().getFullYear().toString() + "/" +
                          (messages[j].getDate().getMonth() + 1).toString() + "/" +
                          messages[j].getDate().getDate().toString();
          // Now update the info in the spreadsheet
          var sheet = load_reporte_sheet(ano, mes.toString());
          
          // But checck first that we haven't save the very same CFD.
          var duplicated = check_for_duplicated(sheet, cfdis[k]);
          if ( duplicated ) {
            // Make as if you haven't touch the message.
            messages[j].markUnread();
            // Acknowledge of the event.
            send_report_message(messages[j], "CFD already saved",
                                "Info", "The CFD " + cfdis[k].serie + cfdis[k].folio +
                                " was previously saved/stored. We will not save it again." +
                                "\n\nMessage details:" +
                                "\n\tFrom:" + messages[j].getFrom() +
                                "\n\tTo: " + messages[j].getTo() +
                                "\n\tSubject: " + messages[j].getSubject());
            continue;
          }

          // Save the blobs.
          var url = "";
          if ( cfdis[k].pdf ) {
            var pdf_file = folder.createFile(cfdis[k].pdf);
            pdf_file.setDescription("CFD. Prov:" + cfdis[k].rfc + ", Folio: " + cfdis[k].folio +
                                    ", Fecha: " + ano + "/" + mes.toString() + "/" + dia);
            url = pdf_file.getUrl();
          }
          var xml_file = folder.createFile(cfdis[k].comprobante);
          xml_file.setDescription("CFD. Prov:" + cfdis[k].rfc + ", Folio: " + cfdis[k].folio +
                                  ", Fecha: " + ano + "/" + mes.toString() + "/" + dia);

          // Now store the information
          sheet.appendRow([ano + "/" + mes.toString() + "/" + dia,
                           cfdis[k].version,
                           cfdis[k].serie,
                           cfdis[k].folio,
                           cfdis[k].uuid,
                           cfdis[k].proveedor,
                           cfdis[k].rfc,
                           cfdi_tipo,
                           cfdis[k].descuento.toString(),
                           cfdis[k].subtotal.toString(),
                           cfdis[k].impuestos.toString(),
                           cfdis[k].total.toString(),
                           recepcion, // Recepcion (del comprobante)
                           "", // Recepcion mercancias
                           "", // Fecha pago
                           "", // Monto pago
                           "", // Observaciones
                           "Vigente", // Estado
                           url, // Url
                           ""]); // Col1

          // Wait for a minute to prevent timeouts
          Utilities.sleep(1000);

          // Apply a label
          var label = get_label_name(cfdis[k].rfc);
          messages[j].getThread().addLabel(GmailApp.getUserLabelByName(label)).markUnread();
          // Archive the message.
          messages[j].getThread().moveToArchive();
        }

      } else {
        // Make as if you haven't touch the message.
        messages[j].markUnread();
      }
    }
  }
}

// Helper function to load the folder in which the CFD
// and its printed version (if any) will be saved.
// In the requested folder doesn't exist, it will be created.
//
// The function requires three arguments
//
//      r   -- The root folder name ("Compras")
//      a   -- The year number as String
//      m   -- The month number as String
//
// The function returns a Folder object pointed
// at the current folder.
function load_folder(message, r, a, m) {
  var folders = null;
  var folder = null;
  var found = false;

  // Load the list of all folders.
  folders = DocsList.getRootFolder().getFolders();

  // Now loop over all the folders in the root directory and try to find the one we want.
  for (var i = 0; i < folders.length; i++) {
    if ( folders[i].getDescription() == r ) {
      found = true;
      break;
    }
  }

  if ( !found ) {
    // Create one and the others beneath
    send_report_message(message, "ROOT Folder (" + r + ") doesn't exist!",
                         "Warning", "Creating folder \"" + r + "\" and all the others needed");
    folder = DocsList.createFolder(r);
    folder.setDescription(r);
    var tmp_folder = folder.createFolder(a);
    tmp_folder.setDescription(a);
    var temp_folder = tmp_folder.createFolder(m);
    temp_folder.setDescription(m);
    folder = DocsList.getFolder(r + "/" + a + "/" + m);
    return folder;
  }

  // If we are here it is because we have at least one folder 'r'.
  // Now start with the year folder
  found = false;
  folders = DocsList.getFolder(r).getFolders();
  for (var i = 0; i < folders.length; i++) {
    if ( folders[i].getDescription() == a ) {
      found = true;
      break;
    }
  }

  if ( !found ) {
    send_report_message(message, "Folder (" + r + "/" + a + ") doesn't exist!",
                        "Warning", "Creating folder \"" + r + "/" + a + "\" and all the others needed.");
    folder = DocsList.getFolder(r).createFolder(a);
    folder.setDescription(a);
    var tmp_folder = folder.createFolder(m);
    tmp_folder.setDescription(m);
    folder = DocsList.getFolder(r + "/" + a + "/" + m);
    return folder;
  }

  // Now the month
  found = false;
  folders = DocsList.getFolder(r + "/" + a).getFolders();
  for (var i = 0; i < folders.length; i++) {
    if ( folders[i].getDescription() == m ) {
      found = true;
      break;
    }
  }

  if ( !found ) {
    send_report_message(message, "Folder (" + r + "/" + a + "/" + m + ") doesn't exist!",
                        "Warning", "Creating folder \"" + r + "/" + a + "/" + m + "\"");
    var folder = DocsList.getFolder(r + "/" + a).createFolder(m);
    folder.setDescription(m);
  }

  folder = DocsList.getFolder(r + "/" + a + "/" + m);
  return folder;
}

// Helper function to verify that no CFD will be
// saved/stored duplicated.
// The parameters to verify that no CFD will be
// duplicated are:
//     - Serie
//     - Folio
//     - Folio fiscal (if any)
//     - R.F.C. del emisor.
//
// The function will return "true" if the CFD is already
// in the system, otherwise will return "false"
//
function check_for_duplicated(sheet, cfd) {
  var found = false;

  var data = sheet.getDataRange().getValues();
  // Skip the header of the sheet
  for (var i = 1; i < data.length; i++) {
    // The columns to search for are: 2, 3, 4 y 6
    if ( data[i][2].toString() === cfd.serie && data[i][3].toString() === cfd.folio &&
         data[i][4].toString() === cfd.uuid  && data[i][6].toString() === cfd.rfc ) {
      found = true;
    }
  }
  // Load all the rows in the sheet.
  return found;
}

// Helper function to load the Spreadsheet
// in which the info will be saved.
// In the case where the spreadsheet doesn't
// exist, we will create one. The list of
// spreadsheets will be also updated.
//
// The function has two arguments:
//
//       a   The full year of the CFD, as a string
//       m   The month number, of the CFD, as a string
//
// the function returns the Sheet object in which we
// will save the CFD data.
//
function load_reporte_sheet(a, m) {

  var base_name = root_name + "-" + a + "-" + m;
  var id = UserProperties.getProperty(base_name + "-key");
  var sheet = undefined;
  if ( id != null ) {
    // Simply load the sheet
    sheet = SpreadsheetApp.openById(id).getActiveSheet();
  } else {
    // Create a sheet and save the parameters to the master ss
    var ss = SpreadsheetApp.create(base_name);
    // Set the table headers.
    ss.appendRow(["Fecha emisión", "Versión", "Serie", "Folio", "Folio fiscal",
                  "Proveedor", "R.F.C", "Tipo", "Descuentos", "Subtotal", "Impuestos",
                  "Total", "Recepción", "Recepción mercancia", "Fecha pago",
                  "Monto pago", "Observaciones", "Estado", "Url", "Col1"]);
    // Make this row looks like a header for the contains.
    ss.getRange("A1:T1").setBackground("black").setFontColor("white");
    id = ss.getId();
    // Save the ss key that you just have created
    UserProperties.setProperty(base_name + "-key", id);
    // Now retrieve the sheet
    sheet = ss.getActiveSheet();
  }

  return sheet;
}

// This function it is meant to seek in the User Properties
// the Gmail mailbox label name under the message will be archived.
//
// To properties (defined by the user) must have as key the
// R.F.C., and its value the Gmail mailbox label. For instance,
// the user may define the property key
//
//           XAXX010101111
//
// with value
//
//           GastosPersonales
//
// In case that label doesn't exist, it will be created.
//
// Aguments:
//            A string object with the R.F.C.
// Returns:
//            A string object with the label name
//            (default 'default_label' variable.)
//
function get_label_name(rfc) {

  // Load the user key value
  var label = UserProperties.getProperty(rfc);
  if ( label === null ) label = default_label;

  // The label already exist as user property. Check
  // if there is a label in the mailbox.
  var labels = GmailApp.getUserLabels();
  var found = false;
  for (var i = 0; i < labels.length; i++) {
    if ( labels[i] === label ) {
      found = true;
      break;
    }
  }
  if ( found != true ) {
    GmailApp.createLabel(label);
  }
  return label;
}

// This function will parse the file attachments
// in the GmailMessage object and try to associate them
// to a CFD/CFDi instances, in case that those attachments
// are so.
//
// The function will also try to save the CFD/CFDi printed
// version only if the file name (without file extension)
// is the same as the CFD/CFDi instance.
//
// Arguments:
//           m         A GmailMessage object
// Returns:
//           An array of CFD/CFDis parsed from the
//           message attachments.
//
function parse_cfdi(m) {

  var facturas = []; // Array of documents.
  var result = []; // The array of all the cfdis

  // Parse all the attachments in the current message
  // First save the CFD instances.
  var att = m.getAttachments();
  var name = null;
  var type = null;
  for (var i = 0; i < att.length; i++) {
    name = att[i].getName();
    type = att[i].getContentType();

    if ( type === "text/xml" || type === "application/xml" ) {

      var documento = {
        pdf_Blob: null,
        xml_Blob: null,
        name:    null
      };
      documento.name = name.substr(0, name.lastIndexOf("."));
      documento.xml_Blob = att[i];
      facturas.push(documento);
    }
  }

  // Now make other loop over all the attachments
  // to associate (if any) all the CFD instances with its printed version
  for (var i = 0; i < att.length; i++) {
    name = att[i].getName();
    type = att[i].getContentType();

    if ( type === "application/pdf" ) {
      // Make a loop over all the entries in the facturas array to associte
      // this printed version with its CFD instance.
      for (var j = 0; j < facturas.length; j++) {
        if ( name == facturas[j].name + ".pdf" || name == facturas[j].name + ".PDF" ) {
          facturas[j].pdf_Blob = att[i];
        }
      }
    }
  }

  // Now, parse all the real invoices (CFD/CFDis)
  for (var i = 0; i < facturas.length; i++) {
    var cfdi = {
      fecha:       null,
      rfc:         null,
      proveedor:   null,
      descuento:   0,
      subtotal:    0,
      impuestos:   0,
      total:       0,
      metodo_pago: "",
      folio:       "",
      serie:       null,
      comprobante: null,
      pdf:         null,
      tipo:        1, // 1 = ingreso, -1 = egreso (declarado en el comprobante)
      version:     null,
      uuid:        null
    };

    var taxes = null;
    var date = null;
    var a = null;
    var doc = Xml.parse(facturas[i].xml_Blob.getDataAsString(), true);

    // Verify that this instance is a CFD/CFDi instance.
    if ( doc.getElement().getName().getLocalName() != "Comprobante" ) continue;

    cfdi.proveedor   = doc.getElement().getElement("Emisor").getAttribute("nombre").getValue();
    cfdi.rfc         = doc.getElement().getElement("Emisor").getAttribute("rfc").getValue();
    cfdi.serie       = doc.getElement().getAttribute("serie").getValue();
    cfdi.folio       = doc.getElement().getAttribute("folio").getValue();
    cfdi.metodo_pago = doc.getElement().getAttribute("metodoDePago").getValue();
    date             = doc.getElement().getAttribute("fecha").getValue();
    a = date.split(/[T:-]/); // Split the date string onto its members.
    cfdi.fecha = new Date(a[0], a[1]-1, a[2], a[3], a[4], a[5]);
    if ( doc.getElement().getAttribute("tipoDeComprobante").getValue().toLowerCase() == "egreso" ) {
      cfdi.tipo = -1;
    }

    // Apply the "tipoDeComprobante" factor.
    cfdi.subtotal   = doc.getElement().getAttribute("subTotal").getValue() * cfdi.tipo;
    cfdi.total      = doc.getElement().getAttribute("total").getValue() * cfdi.tipo;
    // Now, we need to make a loop over all the taxes.
    taxes           = doc.getElement().getElement("Impuestos").getElement("Traslados").getElements();
    for (var j in taxes) {
      cfdi.impuestos += taxes[j].getAttribute("importe").getValue() * cfdi.tipo;
    }

    // Get the "descuento" if any
    if ( doc.getElement().getAttribute("descuento") ) {
      cfdi.descuento = doc.getElement().getAttribute("descuento").getValue();
    } else {
      cfdi.descuento = 0;
    }

    cfdi.version = doc.getElement().getAttribute("version").getValue();

    // Get the "folio fiscal", if any
    if ( parseFloat(cfdi.version) >= 3 ) {
      var timbre = doc.getElement().getElement("Complemento").getElement("TimbreFiscalDigital");
      cfdi.uuid = timbre.getAttribute("UUID").getValue();
    } else {
      cfdi.uuid = "";
    }

    cfdi.comprobante = facturas[i].xml_Blob;
    if ( facturas[i].pdf_Blob ) cfdi.pdf = facturas[i].pdf_Blob;

    result.push(cfdi);
  }

  return result;
}

// Helper function to send to a mailbox
// some error/warning messages
function send_report_message(message, error, level, info) {
  GmailApp.sendEmail(Session.getActiveUser().getUserLoginId(),
                     "UpdateCFDi.gs: " + level + ": " + message.getSubject(), error + "\n" +
                     "Guarda-CFD/CFDi: Automatic message.\n\n" + info);
}
