// File: src/extension.ts
import * as vscode from 'vscode';
import { queryClaude } from './claudeAPI';

export function activate(context) {
    console.log('Claude AI Assistant is now active!');

    let disposable = vscode.commands.registerCommand('vscode-claude-extension.queryClaude', async () => {
        const editor = vscode.window.activeTextEditor;
        if (editor) {
            const selection = editor.selection;
            const text = editor.document.getText(selection);

            if (text) {
                try {
                    const response = await queryClaude(text);
                    editor.edit(editBuilder => {
                        editBuilder.replace(selection, response);
                    });
                } catch (error) {
                    vscode.window.showErrorMessage('Error querying Claude: ' + error.message);
                }
            } else {
                vscode.window.showInformationMessage('Please select some text to query Claude.');
            }
        }
    });

    context.subscriptions.push(disposable);
}

export function deactivate() {}
