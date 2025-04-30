import { getLLMResponse } from "./llm";
import { marked } from 'marked';
import { ChatOpenAI } from "@langchain/openai";

// Function to convert markdown to Word formatting
async function convertMarkdownToWordFormatting(context: Word.RequestContext, markdownText: string) {
    const lines = markdownText.split('\n');
    const selection = context.document.getSelection();
    
    for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        
        // Handle headings
        if (line.startsWith('# ')) {
            const paragraph = selection.insertParagraph(line.substring(2), Word.InsertLocation.before);
            paragraph.style = "Heading 1";
            await context.sync();
        } else if (line.startsWith('## ')) {
            const paragraph = selection.insertParagraph(line.substring(3), Word.InsertLocation.before);
            paragraph.style = "Heading 2";
            await context.sync();
        } else if (line.startsWith('### ')) {
            const paragraph = selection.insertParagraph(line.substring(4), Word.InsertLocation.before);
            paragraph.style = "Heading 3";
            await context.sync();
        } else if (line.startsWith('#### ')) {
            const paragraph = selection.insertParagraph(line.substring(5), Word.InsertLocation.before);
            paragraph.style = "Heading 4";
            await context.sync();
        } else if (line.startsWith('##### ')) {
            const paragraph = selection.insertParagraph(line.substring(6), Word.InsertLocation.before);
            paragraph.style = "Heading 5";
            await context.sync();
        } else if (line.startsWith('###### ')) {
            const paragraph = selection.insertParagraph(line.substring(7), Word.InsertLocation.before);
            paragraph.style = "Heading 6";
            await context.sync();
        } else {
            // Handle normal paragraphs
            const paragraph = selection.insertParagraph(line, Word.InsertLocation.before);
            paragraph.style = "Normal";
            await context.sync();
        }
    }
}

async function concatenatePrompt(promptUrl: string, selectedText: string) {
  if (promptUrl === "") {
    return selectedText;
  } else {
    const response = await fetch(promptUrl);
    if (!response.ok) {
      throw new Error(`Failed to load prompt: ${response.status} ${response.statusText}`);
    }
    const promptText = await response.text();
    return promptText + '\n' + selectedText;
  }
}
export async function askLLMUrlPrompt(userText: string, taskPane: boolean = false, promptUrl: string = "", model: ChatOpenAI) {
  const fullPrompt = await concatenatePrompt(promptUrl, userText);
  console.log('Full prompt:', fullPrompt);
  askLLM(fullPrompt, taskPane, model);
}

export async function askLLMStrPrompt(userText: string, taskPane: boolean = false, promptStr: string = "", model: ChatOpenAI) {
  const fullPrompt = promptStr + "\n" + userText;
  console.log('Full prompt:', fullPrompt);
  askLLM(fullPrompt, taskPane, model);
}


// If taskPane is true, the response is displayed in the taskpane
// If taskPane is false, the response replace the selection
export async function askLLM(prompt: string, taskPane: boolean = false, model: ChatOpenAI) {
    // Show loading overlay
    const overlay = document.getElementById('loadingOverlay');
    if (overlay) {
        overlay.style.display = 'flex';
        overlay.classList.add('active');
    }

    try {
        await Word.run(async (context) => {
            // Get response from LLM
            const llmResponse = await getLLMResponse(prompt, model);
            console.log('Received the LLM response:', llmResponse);

            // Display the response in the taskpane
            if (taskPane) {
              const responseDiv = document.getElementById('response');
              if (responseDiv) {
                  const parsedMarkdown = await marked.parse(llmResponse);
                  console.log('Displaying response in taskpane : \n ' + parsedMarkdown);
                  responseDiv.innerHTML = `<div class="markdown-content">${parsedMarkdown}</div>`;
                  
                  // Scroll to the response section
                  const responseSection = document.querySelector('.response-section');
                  if (responseSection) {
                      responseSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
                  }
              }
            } else {
              // Clear the selection first
              const selection = context.document.getSelection();
              selection.clear();
              
              // Convert markdown to Word formatting
              await convertMarkdownToWordFormatting(context, llmResponse);
            }

            await context.sync();
        });
    } catch (error) {
        console.error('Error in askLLM:', error);
        const responseDiv = document.getElementById('response');
        if (responseDiv) {
            console.log('Displaying error in UI');
            responseDiv.textContent = `An error occurred: ${error.message || error}. Please try again.`;
            
            // Scroll to the response section even on error
            const responseSection = document.querySelector('.response-section');
            if (responseSection) {
                responseSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
            }
        }
    } finally {
        // Hide loading overlay
        if (overlay) {
            overlay.style.display = 'none';
            overlay.classList.remove('active');
        }
    }
    console.log('Finished askLLM function');
}

