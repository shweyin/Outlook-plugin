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
          outputString += '<div class="panel panel-default card-input form-check card"><img src="../../assets/icon-32.png" class="col-2" alt=""/><div class="col-10">';
          outputString += '<h5 class="card-title">' + attachment.name + '</h5>';
          outputString += '<h6 class="card-subtitle mb-2 text-muted">' + attachement.contentType + ' ' + attachment.size + '</h6>';
          outputString += '<p class="card-text">Could add some additional descriptive text somehow........ from parser?</p>';
          outputString += '</div></label></div></div>';
      }
  }

  outputString += '<div class="container"><label>';
  outputString += '<input type="radio" name="attachments" id="html" class="card-input-element" value="item"/>';
  outputString += '<div class="panel panel-default card-input form-check card"><img src="../../assets/icon-32.png" class="col-2" alt=""/><div class="col-10">';
  outputString += '<h5 class="card-title">HTML Message Body</h5>';
  outputString += '<h6 class="card-subtitle mb-2 text-muted">undetermined size</h6>';
  outputString += '<p class="card-text">Could add some additional descriptive text somehow........ from parser?</p>';
  outputString += '</div></label></div></div>';

  document.getElementById("msg-attachments").innerHTML = outputString;
}
