/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */
import { askLLM } from '../lib/word_insertion';

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // Add here all the event listeners
    document.getElementById("chat").onclick = chat;
    document.getElementById("explain").onclick = explain;
    document.getElementById("translateToEnglish").onclick = translateToEnglish;
    document.getElementById("translateToFrench").onclick = translateToFrench;
    document.getElementById("enhance").onclick = enhance;
  }
});

// This function is called when the user clicks the "Chat" button
export async function chat() {
  return Word.run(async (context) => {
    console.log("run button clicked");
    // Get the prompt from the textarea
    const textAreaValue = (document.getElementById('prompt') as HTMLTextAreaElement).value;
    console.log('Retrieved textAreaValue:', textAreaValue);
    await askLLM(textAreaValue, true); 
    await context.sync();
  });
}

// This function is called when the user clicks the "Explain" button
export async function explain() {
  return Word.run(async (context) => {
    // Get the current selection
    const selection = context.document.getSelection();
    selection.load('text');
    await context.sync();
    const selectedText = selection.text;
    await askLLM(selectedText, true, '/prompts/explain.txt');
    await context.sync();
  });
}

// This function is called when the user clicks the "Translate to English" button
export async function translateToEnglish() {
  return Word.run(async (context) => {
    console.log("translate to english button clicked");
    const selection = context.document.getSelection();
    selection.load('text');
    await context.sync();
    const selectedText = selection.text;
    await askLLM(selectedText, false, '/prompts/translateToEnglish.txt');
    await context.sync();
  });
}

// This function is called when the user clicks the "Translate to French" button
export async function translateToFrench() {
  return Word.run(async (context) => {
    console.log("translate to english button clicked");
    const selection = context.document.getSelection();
    selection.load('text');
    await context.sync();
    const selectedText = selection.text;
    await askLLM(selectedText, false, '/prompts/translateToFrench.txt');
    await context.sync();
  });
}

// This function is called when the user clicks the "Enhance" button
export async function enhance() {
  return Word.run(async (context) => {
    console.log("enhance button clicked");
    const selection = context.document.getSelection();
    selection.load('text');
    await context.sync();
    const selectedText = selection.text;
    await askLLM(selectedText, false, '/prompts/enhance.txt');
    await context.sync();
  });
}