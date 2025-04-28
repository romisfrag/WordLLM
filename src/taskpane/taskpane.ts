/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */
import { askLLM } from '../lib/word_insertion';
import { initializeModel } from '../lib/llm';
import { setInLocalStorage, getFromLocalStorage } from '../lib/local_storage';

let model: any = null;

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // Load saved configuration
    const savedBaseURL = getFromLocalStorage('baseURL');
    const savedApiKey = getFromLocalStorage('apiKey');
    
    if (savedBaseURL) {
      (document.getElementById("baseURL") as HTMLInputElement).value = savedBaseURL;
    }
    if (savedApiKey) {
      (document.getElementById("apiKey") as HTMLInputElement).value = savedApiKey;
    }

    // Initialize model with saved or default values
    const baseURL = (document.getElementById("baseURL") as HTMLInputElement).value;
    const apiKey = (document.getElementById("apiKey") as HTMLInputElement).value;
    model = initializeModel(baseURL, apiKey);

    // Add event listeners for configuration changes
    document.getElementById("baseURL").addEventListener("change", updateModel);
    document.getElementById("apiKey").addEventListener("change", updateModel);
    document.getElementById("saveConfig").addEventListener("click", saveConfiguration);

    // Add here all the event listeners
    document.getElementById("chat").onclick = chat;
    document.getElementById("explain").onclick = explain;
    document.getElementById("translateToEnglish").onclick = translateToEnglish;
    document.getElementById("translateToFrench").onclick = translateToFrench;
    document.getElementById("enhance").onclick = enhance;
  }
});

function updateModel() {
  const baseURL = (document.getElementById("baseURL") as HTMLInputElement).value;
  const apiKey = (document.getElementById("apiKey") as HTMLInputElement).value;
  model = initializeModel(baseURL, apiKey);
}

function saveConfiguration() {
  const baseURL = (document.getElementById("baseURL") as HTMLInputElement).value;
  const apiKey = (document.getElementById("apiKey") as HTMLInputElement).value;
  
  setInLocalStorage('baseURL', baseURL);
  setInLocalStorage('apiKey', apiKey);
  
  // Show a success message
  const responseDiv = document.getElementById('response');
  if (responseDiv) {
    responseDiv.innerHTML = '<div class="markdown-content">Configuration saved successfully!</div>';
  }
}

// This function is called when the user clicks the "Chat" button
export async function chat() {
  // Get the prompt from the textarea
  const textAreaValue = (document.getElementById('prompt') as HTMLTextAreaElement).value;
  console.log('Retrieved textAreaValue:', textAreaValue);
  await askLLM(textAreaValue, true, "", model);
}

// This function is called when the user clicks the "Explain" button
export async function explain() {
  return Word.run(async (context) => {
    // Get the current selection
    const selection = context.document.getSelection();
    selection.load('text');
    await context.sync();
    const selectedText = selection.text;
    await askLLM(selectedText, true, '/prompts/explain.txt', model);
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
    await askLLM(selectedText, false, '/prompts/translateToEnglish.txt', model);
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
    await askLLM(selectedText, false, '/prompts/translateToFrench.txt', model);
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
    await askLLM(selectedText, false, '/prompts/enhance.txt', model);
    await context.sync();
  });
}