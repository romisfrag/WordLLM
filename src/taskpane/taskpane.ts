/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */
import { askLLMUrlPrompt, askLLMStrPrompt } from '../lib/word_insertion';
import { initializeModel, fetchAvailableModels, filterModels } from '../lib/llm';
import { setInLocalStorage, getFromLocalStorage } from '../lib/local_storage';

let model: any = null;
let availableModels: string[] = [];
let currentPromptType: 'replaceSelection' | 'taskpane' | null = null;

// Loading overlay functions
function showLoadingOverlay() {
    console.log('Showing loading overlay');
    const overlay = document.getElementById('loadingOverlay');
    if (overlay) {
        overlay.style.display = 'flex';
        overlay.classList.add('active');
    } else {
        console.error('Loading overlay element not found');
    }
}

function hideLoadingOverlay() {
    console.log('Hiding loading overlay');
    const overlay = document.getElementById('loadingOverlay');
    if (overlay) {
        overlay.style.display = 'none';
        overlay.classList.remove('active');
    }
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // Load saved configuration
    const savedBaseURL = getFromLocalStorage('baseURL');
    const savedApiKey = getFromLocalStorage('apiKey');
    const savedModel = getFromLocalStorage('selectedModel');
    
    if (savedBaseURL) {
      (document.getElementById("baseURL") as HTMLInputElement).value = savedBaseURL;
    }
    if (savedApiKey) {
      (document.getElementById("apiKey") as HTMLInputElement).value = savedApiKey;
    }

    // Initialize model with saved or default values
    const baseURL = (document.getElementById("baseURL") as HTMLInputElement).value;
    const apiKey = (document.getElementById("apiKey") as HTMLInputElement).value;
    const selectedModel = savedModel || undefined;
    model = initializeModel(baseURL, apiKey, selectedModel);

    // Add event listeners for configuration changes
    document.getElementById("baseURL").addEventListener("change", updateModel);
    document.getElementById("apiKey").addEventListener("change", updateModel);
    document.getElementById("saveConfig").addEventListener("click", saveConfiguration);
    document.getElementById("modelSearch").addEventListener("input", handleModelSearch);
    document.getElementById("modelSelect").addEventListener("change", handleModelChange);
    document.getElementById("toggleConfig").addEventListener("click", toggleConfigSection);

    // Add tab switching functionality
    const tabButtons = document.querySelectorAll('.tab-button');
    tabButtons.forEach(button => {
      button.addEventListener('click', () => {
        // Remove active class from all buttons and content
        tabButtons.forEach(btn => btn.classList.remove('active'));
        document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));
        
        // Add active class to clicked button and corresponding content
        button.classList.add('active');
        const tabId = button.getAttribute('data-tab');
        document.getElementById(`${tabId}-tab`).classList.add('active');
      });
    });

    // Add event listeners for dev mode execute buttons
    document.getElementById("executeReplaceSelection").addEventListener("click", executeReplaceSelection);
    document.getElementById("executeReplyTaskpane").addEventListener("click", executeReplyTaskpane);
    document.getElementById("saveReplaceSelectionPrompt").addEventListener("click", () => showPromptNamePopup('replaceSelection'));
    document.getElementById("saveTaskpanePrompt").addEventListener("click", () => showPromptNamePopup('taskpane'));
    document.getElementById("savePromptConfirm").addEventListener("click", savePromptWithName);
    document.getElementById("cancelPromptSave").addEventListener("click", hidePromptNamePopup);

    // Load saved custom prompts
    loadCustomPrompts();

    // Fetch available models
    fetchModels();

    // Add here all the event listeners
    document.getElementById("chat").onclick = chat;
    document.getElementById("explain").onclick = explain;
    document.getElementById("translateToEnglish").onclick = translateToEnglish;
    document.getElementById("translateToFrench").onclick = translateToFrench;
    document.getElementById("enhance").onclick = enhance;
  }
});

async function fetchModels() {
    const modelSelect = document.getElementById("modelSelect") as HTMLSelectElement;
    modelSelect.classList.add("loading");
    
    try {
        const baseURL = (document.getElementById("baseURL") as HTMLInputElement).value;
        const apiKey = (document.getElementById("apiKey") as HTMLInputElement).value;
        
        availableModels = await fetchAvailableModels(baseURL, apiKey);
        updateModelDropdown(availableModels);
        
        // Restore saved model selection if exists
        const savedModel = getFromLocalStorage('selectedModel');
        if (savedModel && availableModels.includes(savedModel)) {
            modelSelect.value = savedModel;
        }
    } catch (error) {
        console.error("Error fetching models:", error);
        modelSelect.innerHTML = '<option value="">Error loading models</option>';
    } finally {
        modelSelect.classList.remove("loading");
    }
}

function updateModelDropdown(models: string[]) {
    const modelSelect = document.getElementById("modelSelect") as HTMLSelectElement;
    const searchTerm = (document.getElementById("modelSearch") as HTMLInputElement).value;
    
    // Filter models based on search term
    const filteredModels = filterModels(models, searchTerm);
    
    // Update dropdown options
    modelSelect.innerHTML = filteredModels.length > 0 
        ? filteredModels.map(model => `<option value="${model}">${model}</option>`).join('')
        : '<option value="">No models found</option>';
    
    // If there's only one model after filtering, select it and update the model
    if (filteredModels.length === 1) {
        modelSelect.value = filteredModels[0];
        updateModel();
    }
}

function handleModelSearch() {
    updateModelDropdown(availableModels);
}

function handleModelChange() {
    const selectedModel = (document.getElementById("modelSelect") as HTMLSelectElement).value;
    if (selectedModel) {
        setInLocalStorage('selectedModel', selectedModel);
        updateModel();
    }
}

function updateModel() {
    const baseURL = (document.getElementById("baseURL") as HTMLInputElement).value;
    const apiKey = (document.getElementById("apiKey") as HTMLInputElement).value;
    const selectedModel = (document.getElementById("modelSelect") as HTMLSelectElement).value;
    
    // Force a new model instance to be created
    model = initializeModel(baseURL, apiKey, selectedModel);
    console.log('Model updated to:', selectedModel);
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

  // Fetch models with the new configuration
  fetchModels();
}

// This function is called when the user clicks the "Chat" button
export async function chat() {
    // Get the prompt from the textarea
    const textAreaValue = (document.getElementById('prompt') as HTMLTextAreaElement).value;
    console.log('Retrieved textAreaValue:', textAreaValue);
    await askLLMUrlPrompt(textAreaValue, true, "", model);
}

// This function is called when the user clicks the "Explain" button
export async function explain() {
    return Word.run(async (context) => {
        // Get the current selection
        const selection = context.document.getSelection();
        selection.load('text');
        await context.sync();
        const selectedText = selection.text;
        await askLLMUrlPrompt(selectedText, true, '/prompts/explain.txt', model);
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
        await askLLMUrlPrompt(selectedText, false, '/prompts/translateToEnglish.txt', model);
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
        await askLLMUrlPrompt(selectedText, false, '/prompts/translateToFrench.txt', model);
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
        await askLLMUrlPrompt(selectedText, false, '/prompts/enhance.txt', model);
        await context.sync();
    });
}

// Dev Mode Functions
async function executeReplaceSelection() {
    const prompt = (document.getElementById("promptReplaceSelection") as HTMLTextAreaElement).value;
    if (!prompt) return;

    return Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.load('text');
        await context.sync();
        const selectedText = selection.text;
        await askLLMStrPrompt(selectedText, false, prompt, model);
        await context.sync();
    });
}

async function executeReplyTaskpane() {
    const prompt = (document.getElementById("promptReplyTaskpane") as HTMLTextAreaElement).value;
    if (!prompt) return;

    return Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.load('text');
        await context.sync();
        const selectedText = selection.text;
        await askLLMStrPrompt(selectedText, true, prompt, model);
        await context.sync();
    });
}

function showPromptNamePopup(promptType: 'replaceSelection' | 'taskpane') {
    currentPromptType = promptType;
    const popup = document.getElementById('promptNamePopup');
    if (popup) {
        popup.classList.add('active');
        // Focus the input field
        const input = document.getElementById('promptName') as HTMLInputElement;
        if (input) {
            input.value = '';
            input.focus();
        }
    }
}

function hidePromptNamePopup() {
    const popup = document.getElementById('promptNamePopup');
    if (popup) {
        popup.classList.remove('active');
    }
    currentPromptType = null;
}

function savePromptWithName() {
    if (!currentPromptType) return;

    const promptName = (document.getElementById('promptName') as HTMLInputElement).value.trim();
    if (!promptName) {
        showSuccessMessage('Please enter a name for the prompt');
        return;
    }

    let promptText = '';
    if (currentPromptType === 'replaceSelection') {
        promptText = (document.getElementById("promptReplaceSelection") as HTMLTextAreaElement).value;
    } else {
        promptText = (document.getElementById("promptReplyTaskpane") as HTMLTextAreaElement).value;
    }

    if (promptText) {
        // Save the prompt with its name
        const savedPrompts = JSON.parse(getFromLocalStorage('savedPrompts') || '{}');
        savedPrompts[promptName] = {
            type: currentPromptType,
            text: promptText
        };
        setInLocalStorage('savedPrompts', JSON.stringify(savedPrompts));
        showSuccessMessage(`Prompt "${promptName}" saved successfully!`);
        hidePromptNamePopup();
        loadCustomPrompts(); // Reload the custom prompts list
    }
}

function showSuccessMessage(message: string) {
    const responseDiv = document.getElementById('response');
    if (responseDiv) {
        responseDiv.innerHTML = `<div class="markdown-content">${message}</div>`;
    }
}

function loadCustomPrompts() {
    const savedPrompts = JSON.parse(getFromLocalStorage('savedPrompts') || '{}');
    const customReplacePrompts = document.getElementById('customReplacePrompts');
    const customTaskpanePrompts = document.getElementById('customTaskpanePrompts');

    if (customReplacePrompts && customTaskpanePrompts) {
        customReplacePrompts.innerHTML = '';
        customTaskpanePrompts.innerHTML = '';

        Object.entries(savedPrompts).forEach(([name, promptData]: [string, any]) => {
            const button = createCustomPromptButton(name, promptData);
            if (promptData.type === 'replaceSelection') {
                customReplacePrompts.appendChild(button);
            } else {
                customTaskpanePrompts.appendChild(button);
            }
        });
    }
}

function createCustomPromptButton(name: string, promptData: any) {
    const button = document.createElement('button');
    button.className = 'action-button custom-prompt-button';
    button.innerHTML = `
        <span class="button-text">${name}</span>
        <button class="delete-prompt-button" title="Delete prompt">🗑️</button>
    `;

    button.addEventListener('click', () => {
        if (promptData.type === 'replaceSelection') {
            return Word.run(async (context) => {
                const selection = context.document.getSelection();
                selection.load('text');
                await context.sync();
                const selectedText = selection.text;
                await askLLMStrPrompt(selectedText, false, promptData.text, model);
                await context.sync();
            });
        } else {
            return Word.run(async (context) => {
                const selection = context.document.getSelection();
                selection.load('text');
                await context.sync();
                const selectedText = selection.text;
                await askLLMStrPrompt(selectedText, true, promptData.text, model);
                await context.sync();
            });
        }
    });

    const deleteButton = button.querySelector('.delete-prompt-button');
    if (deleteButton) {
        deleteButton.addEventListener('click', (e) => {
            e.stopPropagation();
            deleteCustomPrompt(name);
        });
    }

    return button;
}

function deleteCustomPrompt(name: string) {
    const savedPrompts = JSON.parse(getFromLocalStorage('savedPrompts') || '{}');
    delete savedPrompts[name];
    setInLocalStorage('savedPrompts', JSON.stringify(savedPrompts));
    loadCustomPrompts();
    showSuccessMessage(`Prompt "${name}" deleted successfully!`);
}

function toggleConfigSection() {
    const configSection = document.querySelector('.config-section');
    if (configSection) {
        configSection.classList.toggle('collapsed');
        const button = document.getElementById('toggleConfig');
        if (button) {
            button.innerHTML = configSection.classList.contains('collapsed') 
                ? '<span class="button-icon">+</span>' 
                : '<span class="button-icon">−</span>';
        }
    }
}