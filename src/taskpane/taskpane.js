/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

/* global document, Office */

Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    //document.getElementById("mySubmit").onclick = mySubmit;
    document.getElementById("my-form").onsubmit = mySubmit;
  }
});

export async function run() {
  /**
   * Insert your Outlook code here
   */
  // Get a reference to the current message
  var item = Office.context.mailbox.item;
  // Write message property value to the task pane
  document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + item.subject;
  
  // Generate attachment list
  var outputString = "";

  if (item.attachments.length > 0) {
      for (var i = 0 ; i < item.attachments.length ; i++) {
          var attachment = item.attachments[i];

          outputString += '<div class="container"><label>';
          outputString += '<input type="radio" name="attachments" id="attachment' + i + '" class="card-input-element" value="' + attachment.id + '"/>';
          outputString += '<div class="panel panel-default card-input form-check card" id="msgCont"><div class="row"><img class="col-2" width="50" height="50" float-left" src="../../assets/icon-32.png"  alt=""/><div class="col-10">';
          outputString += '<h5 class="card-title">' + attachment.name + '</h5>';
          outputString += '<h6 class="card-subtitle mb-2 text-muted">' + attachment.contentType + ' ' + attachment.size + '</h6>';
          outputString += '<p class="card-text">Could add some additional descriptive text somehow........ from parser?</p>';
          outputString += '</div></div></div></label></div>';
      }
  }

  outputString += '<div class="container"><label>';
  outputString += '<input type="radio" name="attachments" id="html" class="card-input-element" value="item"/>';
  outputString += '<div class="panel panel-default card-input form-check card" id="msgCont"><div class="row"><img src="../../assets/icon-32.png" class="col-2" width="50" height="50" alt=""/><div class="col-10">';
  outputString += '<h5 class="card-title">HTML Message Body</h5>';
  outputString += '<h6 class="card-subtitle mb-2 text-muted">undetermined size</h6>';
  outputString += '<p class="card-text">Could add some additional descriptive text somehow........ from parser?</p>';
  outputString += '</div></div></div></label></div>';

  document.getElementById("msg-attachments").innerHTML = outputString;

}

export async function mySubmit(event) {
  event.preventDefault();
  var form = document.getElementById("my-form");
  var attachments = form.elements["attachments"];
  var output ="";
  var attachmentId = '';

  output += document.getElementById("company-name").value.toUpperCase() + "-";
  output += document.getElementById("product-name").value.toUpperCase() + "-";
  output += document.getElementById("date").value.replace(/-/g, '');
  output += document.getElementById("invoice-num").value.toUpperCase(); 

  // Get the selected attachment ID
  for (var i = 0; i < attachments.length; i++) {
    if (attachments[i].checked) {
      attachmentId = attachments[i].value;
      i = attachments.length;
    }
  }
  
  // Get a reference to the current message
  var item = Office.context.mailbox.item;

  if (attachmentId != 'html')
    item.getAttachmentContentAsync(attachmentId, {asyncContext: {item: item, filename: output, oldId: attachmentId}}, fileAttachment);
  else
    fileHTMLAttachment(item);
}

function fileAttachment (attachment) {
  attachment.asyncContext.item.addFileAttachmentFromBase64Async(attachment.value.content, attachment.asyncContext.filename, 
    { 
      asyncContext: 
      { 
        item: attachment.asyncContext.item, 
        oldId: attachment.asyncContext.oldId
      }
    }, removeAttachment);
}

function removeAttachment (result) {
  var email = new Office.MessageCompose();
  document.getElementById("test-output").innerHTML = result.asyncContext.oldId;
  result.asyncContext.item.removeAttachmentAsync(result.asyncContext.oldId);
  run();
}

function fileHTMLAttachment (item) {
  //console.log("Not yet built. Item: " + item);
}