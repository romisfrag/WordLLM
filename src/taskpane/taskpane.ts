/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */
import { sendPrompt } from '../lib/word_insertion';

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // Add here all the event listeners
    document.getElementById("run").onclick = run;
  }
});

// This function is called when the user clicks the "Run" button
export async function run() {
  return Word.run(async (context) => {
    console.log("run button clicked");
    await sendPrompt(); 
    await context.sync();
  });
}
