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
    document.getElementById("mySubmit").onclick = mySubmit;
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
  
  var outputString = "";

  if (item.attachments.length > 0) { 
    for (var i = 0 ; i < item.attachments.length ; i++) { 
      var attachment = item.attachments[i]; 
      outputString += "<BR>" + i + ". Name: "; 
      outputString += attachment.name; 
      // outputString += "<BR>ID: " + attachment.id; 
      outputString += "<BR>contentType: " + attachment.contentType; 
      outputString += "<BR>size: " + attachment.size; 
      outputString += "<BR>attachmentType: " + attachment.attachmentType; 
      outputString += "<BR>isInline: " + attachment.isInline; 
    } 
  }

  document.getElementById("test-output").innerHTML = outputString;

}

export async function mySubmit() {
  var output ="";

  output += document.getElementById("company-name").value.toUpperCase() + "-";
  output += document.getElementById("product-name").value.toUpperCase() + "-";
  output += document.getElementById("date").value.replace(/-/g, '');
  output += document.getElementById("invoice-num").value.toUpperCase() + "-"; 


  document.getElementById("test-output").innerHTML = output;
}