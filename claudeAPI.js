import axios from 'axios';
import * as vscode from 'vscode';

const API_URL = 'https://api.anthropic.com/v1/messages';

export async function queryClaude(prompt) {
    const config = vscode.workspace.getConfiguration('claudeAssistant');
    const apiKey = config.get('apiKey');

    if (!apiKey) throw new Error('Claude API key is not set. Please set it in the extension settings.');
    
    try {
        const response = await axios.post(API_URL, {
            messages: [{ role: 'user', content: prompt }],
            model: 'claude-3-opus-20240229',
            max_tokens: 1000
        }, {
            headers: {
                'Content-Type': 'application/json',
                'X-API-Key': apiKey
            }
        });

        return response.data.content[0].text;
    } catch (error) {
        console.error('Error querying Claude:', error);
        throw new Error('Failed to get response from Claude');
    }
}
