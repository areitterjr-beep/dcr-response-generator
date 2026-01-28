// DCR Response Generator - Main Application Logic

// DOM Elements
const settingsToggle = document.getElementById('settingsToggle');
const settingsContent = document.getElementById('settingsContent');
const toggleArrow = document.getElementById('toggleArrow');
const saveSettingsBtn = document.getElementById('saveSettings');
const saveIndicator = document.getElementById('saveIndicator');

const adoOrgInput = document.getElementById('adoOrg');
const adoPatInput = document.getElementById('adoPat');
const aoaiEndpointInput = document.getElementById('aoaiEndpoint');
const aoaiDeploymentInput = document.getElementById('aoaiDeployment');
const aoaiTokenInput = document.getElementById('aoaiToken');
const refreshTokenBtn = document.getElementById('refreshTokenBtn');

const workItemUrlInput = document.getElementById('workItemUrl');
const fetchBtn = document.getElementById('fetchBtn');
const dropZone = document.getElementById('dropZone');

const previewSection = document.getElementById('previewSection');
const wiType = document.getElementById('wiType');
const wiId = document.getElementById('wiId');
const wiTitle = document.getElementById('wiTitle');
const wiState = document.getElementById('wiState');
const wiArea = document.getElementById('wiArea');
const wiAssigned = document.getElementById('wiAssigned');
const wiDescContent = document.getElementById('wiDescContent');

const generateBtn = document.getElementById('generateBtn');
const additionalContextInput = document.getElementById('additionalContext');
const responseSection = document.getElementById('responseSection');
const responseText = document.getElementById('responseText');
const regenerateBtn = document.getElementById('regenerateBtn');
const copyBtn = document.getElementById('copyBtn');

const statusBar = document.getElementById('statusBar');
const statusMessage = document.getElementById('statusMessage');

// State
let currentWorkItem = null;

// Initialize
document.addEventListener('DOMContentLoaded', () => {
    loadSettings();
    setupEventListeners();
});

function setupEventListeners() {
    // Settings toggle
    settingsToggle.addEventListener('click', () => {
        settingsContent.classList.toggle('open');
        toggleArrow.classList.toggle('open');
    });

    // Save settings
    saveSettingsBtn.addEventListener('click', saveSettings);

    // Fetch work item
    fetchBtn.addEventListener('click', () => fetchWorkItem());
    workItemUrlInput.addEventListener('keypress', (e) => {
        if (e.key === 'Enter') fetchWorkItem();
    });

    // Drag and drop
    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('drag-over');
    });

    dropZone.addEventListener('dragleave', () => {
        dropZone.classList.remove('drag-over');
    });

    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('drag-over');
        const text = e.dataTransfer.getData('text');
        if (text && text.includes('visualstudio.com')) {
            workItemUrlInput.value = text;
            fetchWorkItem();
        }
    });

    // Generate response
    generateBtn.addEventListener('click', () => generateResponse());
    regenerateBtn.addEventListener('click', () => generateResponse());

    // Copy to clipboard
    copyBtn.addEventListener('click', copyToClipboard);

    // Token refresh button
    refreshTokenBtn.addEventListener('click', updateTokenStatus);

    // Update token status when token field changes
    aoaiTokenInput.addEventListener('input', updateTokenStatus);
}

// Settings Management
function loadSettings() {
    const settings = JSON.parse(localStorage.getItem('dcrGenSettings') || '{}');
    adoOrgInput.value = settings.adoOrg || 'https://office.visualstudio.com';
    adoPatInput.value = settings.adoPat || '';
    aoaiEndpointInput.value = settings.aoaiEndpoint || 'https://augloop-cs-test-eastus-shared-open-ai-0.openai.azure.com';
    aoaiDeploymentInput.value = settings.aoaiDeployment || 'gpt-4o';
    aoaiTokenInput.value = settings.aoaiToken || '';

    // Show settings if not configured
    if (!settings.adoPat || !settings.aoaiToken) {
        settingsContent.classList.add('open');
        toggleArrow.classList.add('open');
    }

    // Update token status button
    updateTokenStatus();
}

function saveSettings() {
    const settings = {
        adoOrg: adoOrgInput.value.trim(),
        adoPat: adoPatInput.value.trim(),
        aoaiEndpoint: aoaiEndpointInput.value.trim(),
        aoaiDeployment: aoaiDeploymentInput.value.trim(),
        aoaiToken: aoaiTokenInput.value.replace(/\s/g, '') // Remove ALL whitespace from token
    };
    localStorage.setItem('dcrGenSettings', JSON.stringify(settings));
    saveIndicator.textContent = 'Saved!';
    setTimeout(() => saveIndicator.textContent = '', 2000);
    updateTokenStatus();
}

// Check if token is a valid JWT and show expiration status
function updateTokenStatus() {
    const token = aoaiTokenInput.value.trim();
    if (!token) {
        refreshTokenBtn.textContent = 'No token set - get one from Azure CLI';
        refreshTokenBtn.style.color = '#dc3545';
        return;
    }

    try {
        // JWT tokens have 3 parts separated by dots
        const parts = token.split('.');
        if (parts.length !== 3) {
            refreshTokenBtn.textContent = 'Invalid token format';
            refreshTokenBtn.style.color = '#dc3545';
            return;
        }

        // Decode the payload (second part)
        const payload = JSON.parse(atob(parts[1]));
        const expTime = payload.exp * 1000; // Convert to milliseconds
        const now = Date.now();
        const timeLeft = expTime - now;

        if (timeLeft <= 0) {
            refreshTokenBtn.textContent = 'Token expired! Get a new one from Azure CLI';
            refreshTokenBtn.style.color = '#dc3545';
        } else if (timeLeft < 5 * 60 * 1000) { // Less than 5 minutes
            refreshTokenBtn.textContent = 'Token expires in less than 5 minutes - refresh soon!';
            refreshTokenBtn.style.color = '#ffc107';
        } else {
            const minutesLeft = Math.floor(timeLeft / 60000);
            refreshTokenBtn.textContent = `Token valid - expires in ${minutesLeft} minutes`;
            refreshTokenBtn.style.color = '#28a745';
        }
    } catch (e) {
        refreshTokenBtn.textContent = 'Could not parse token - may still work';
        refreshTokenBtn.style.color = '#6c757d';
    }
}

function getSettings() {
    return JSON.parse(localStorage.getItem('dcrGenSettings') || '{}');
}

// Status Messages
function showStatus(message, type = 'loading') {
    statusMessage.innerHTML = type === 'loading'
        ? `<span class="loading"></span>${message}`
        : message;
    statusBar.className = 'status-bar visible ' + type;
}

function hideStatus() {
    statusBar.classList.remove('visible');
}

// Parse Work Item URL
function parseWorkItemUrl(url) {
    // Patterns:
    // https://office.visualstudio.com/OC/_workitems/edit/12345
    // https://dev.azure.com/org/project/_workitems/edit/12345
    const patterns = [
        /https:\/\/([^\/]+)\/([^\/]+)\/_workitems\/edit\/(\d+)/,
        /https:\/\/dev\.azure\.com\/([^\/]+)\/([^\/]+)\/_workitems\/edit\/(\d+)/
    ];

    for (const pattern of patterns) {
        const match = url.match(pattern);
        if (match) {
            return {
                org: match[1].includes('dev.azure.com') ? match[1] : `https://${match[1]}`,
                project: match[2],
                id: match[3]
            };
        }
    }
    return null;
}

// Fetch Work Item from Azure DevOps
async function fetchWorkItem() {
    const url = workItemUrlInput.value.trim();
    if (!url) {
        showStatus('Please enter a work item URL', 'error');
        setTimeout(hideStatus, 3000);
        return;
    }

    const parsed = parseWorkItemUrl(url);
    if (!parsed) {
        showStatus('Invalid work item URL format', 'error');
        setTimeout(hideStatus, 3000);
        return;
    }

    const settings = getSettings();
    if (!settings.adoPat) {
        showStatus('Please configure your Azure DevOps PAT in Settings', 'error');
        setTimeout(hideStatus, 3000);
        settingsContent.classList.add('open');
        toggleArrow.classList.add('open');
        return;
    }

    showStatus('Fetching work item...');

    try {
        // Determine the base URL
        let baseUrl = settings.adoOrg;
        if (!baseUrl.startsWith('http')) {
            baseUrl = 'https://' + baseUrl;
        }

        const apiUrl = `${baseUrl}/${parsed.project}/_apis/wit/workitems/${parsed.id}?$expand=all&api-version=7.0`;

        const response = await fetch(apiUrl, {
            headers: {
                'Authorization': 'Basic ' + btoa(':' + settings.adoPat),
                'Content-Type': 'application/json'
            }
        });

        if (!response.ok) {
            throw new Error(`HTTP ${response.status}: ${response.statusText}`);
        }

        const data = await response.json();
        currentWorkItem = data;
        displayWorkItem(data);
        showStatus('Work item loaded successfully', 'success');
        setTimeout(hideStatus, 2000);

    } catch (error) {
        console.error('Fetch error:', error);
        showStatus(`Failed to fetch work item: ${error.message}`, 'error');
        setTimeout(hideStatus, 5000);
    }
}

// Display Work Item
function displayWorkItem(workItem) {
    const fields = workItem.fields;

    // Type
    const type = fields['System.WorkItemType'] || 'Bug';
    wiType.textContent = type;
    wiType.className = 'work-item-type ' + type.toLowerCase();

    // ID
    wiId.textContent = '#' + workItem.id;

    // Title
    wiTitle.textContent = fields['System.Title'] || 'Untitled';

    // Meta
    wiState.textContent = fields['System.State'] || '-';
    wiArea.textContent = fields['System.AreaPath'] || '-';
    wiAssigned.textContent = fields['System.AssignedTo']?.displayName || 'Unassigned';

    // Description - strip HTML tags for display
    let description = fields['System.Description'] || 'No description provided.';
    // Create a temporary element to parse HTML and extract text
    const tempDiv = document.createElement('div');
    tempDiv.innerHTML = description;
    wiDescContent.textContent = tempDiv.textContent || tempDiv.innerText || description;

    // Show preview section
    previewSection.classList.add('visible');
}

// Generate Response using Azure OpenAI
async function generateResponse() {
    if (!currentWorkItem) {
        showStatus('Please fetch a work item first', 'error');
        setTimeout(hideStatus, 3000);
        return;
    }

    const settings = getSettings();
    if (!settings.aoaiEndpoint || !settings.aoaiToken) {
        showStatus('Please configure Azure OpenAI endpoint and token', 'error');
        setTimeout(hideStatus, 3000);
        settingsContent.classList.add('open');
        toggleArrow.classList.add('open');
        return;
    }

    showStatus('Generating response...');
    generateBtn.disabled = true;

    try {
        const fields = currentWorkItem.fields;
        const title = fields['System.Title'] || 'Untitled DCR';
        const description = fields['System.Description'] || 'No description provided.';

        // Build the prompt
        const systemPrompt = `You are a professional technical writer for Microsoft. Your task is to write customer-friendly rejection responses for Design Change Requests (DCRs).

IMPORTANT FORMAT REQUIREMENTS:
- Start the response with "Customer Friendly Rejection" on its own line, followed by a blank line
- Do NOT include any salutation like "Dear [Customer]" or "Dear [Name]"
- Do NOT include any sign-off like "Warm regards", "Best regards", signature blocks, or names at the end
- Jump straight into the content after the header

The response content should follow this structure:
1. Opening: Thank the customer for submitting the DCR and acknowledge the specific request and business context they provided.
2. Rejection: Clearly but kindly state that we cannot proceed with this change at this time. Provide a technical or business reason (e.g., significant cross-platform changes required, current roadmap priorities focused on reliability/performance/cross-platform consistency, etc.).
3. Workaround: ONLY include a workaround section if a specific workaround is explicitly mentioned in the work item description or additional context. Do NOT invent or suggest workarounds on your own. If no workaround is provided, skip this section entirely.
4. Future consideration: Mention the request has been added to the backlog for future consideration and will be evaluated in upcoming planning cycles. Note that while you cannot provide a timeline, the request will be kept active.
5. Closing: Express understanding of the impact and appreciation for the customer's partnership. Invite them to reach out if they have additional questions.

Keep the tone professional, empathetic, and constructive. Do not use excessive corporate jargon. Be concise but thorough.`;

        const additionalContext = additionalContextInput.value.trim();

        let userPrompt = `Please write a rejection response for the following Design Change Request:

Title: ${title}

Description/Details:
${stripHtml(description)}`;

        if (additionalContext) {
            userPrompt += `

Additional Context from Support Engineer:
${additionalContext}`;
        }

        userPrompt += `

Generate a complete, ready-to-send response following the structure provided. Make sure to:
- Start with "Customer Friendly Rejection" header followed by a blank line
- Do NOT include any "Dear..." salutation or "Warm regards" sign-off
- Reference the specific feature/change they requested
- Acknowledge their business context if mentioned
- ONLY mention a workaround if one is explicitly stated in the description or additional context above - do NOT make up workarounds
- Keep the response professional but warm`;

        // Use proxy to bypass CORS when not on localhost
        const isLocalhost = window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1';
        const proxyUrl = 'https://dcr-proxy.a-reitterjr.workers.dev';

        let response;
        if (isLocalhost) {
            // Direct call for local development
            const apiUrl = `${settings.aoaiEndpoint}/openai/deployments/${settings.aoaiDeployment}/chat/completions?api-version=2024-02-15-preview`;
            response = await fetch(apiUrl, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': `Bearer ${settings.aoaiToken}`
                },
                body: JSON.stringify({
                    messages: [
                        { role: 'system', content: systemPrompt },
                        { role: 'user', content: userPrompt }
                    ],
                    max_tokens: 1500,
                    temperature: 0.7
                })
            });
        } else {
            // Use proxy for deployed version
            response = await fetch(proxyUrl, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({
                    endpoint: settings.aoaiEndpoint,
                    deployment: settings.aoaiDeployment,
                    token: settings.aoaiToken,
                    messages: [
                        { role: 'system', content: systemPrompt },
                        { role: 'user', content: userPrompt }
                    ],
                    max_tokens: 1500,
                    temperature: 0.7
                })
            });
        }

        if (!response.ok) {
            const errorData = await response.json().catch(() => ({}));
            throw new Error(errorData.error?.message || `HTTP ${response.status}`);
        }

        const data = await response.json();
        const generatedText = data.choices[0]?.message?.content || 'No response generated.';

        responseText.value = generatedText;
        responseSection.classList.add('visible');

        showStatus('Response generated successfully', 'success');
        setTimeout(hideStatus, 2000);

    } catch (error) {
        console.error('Generation error:', error);
        showStatus(`Failed to generate response: ${error.message}`, 'error');
        setTimeout(hideStatus, 5000);
    } finally {
        generateBtn.disabled = false;
    }
}

// Helper: Strip HTML tags
function stripHtml(html) {
    const tempDiv = document.createElement('div');
    tempDiv.innerHTML = html;
    return tempDiv.textContent || tempDiv.innerText || '';
}

// Copy to Clipboard
async function copyToClipboard() {
    const text = responseText.value;
    if (!text) {
        showStatus('Nothing to copy', 'error');
        setTimeout(hideStatus, 2000);
        return;
    }

    try {
        await navigator.clipboard.writeText(text);
        showStatus('Copied to clipboard!', 'success');
        setTimeout(hideStatus, 2000);
    } catch (error) {
        // Fallback for older browsers
        responseText.select();
        document.execCommand('copy');
        showStatus('Copied to clipboard!', 'success');
        setTimeout(hideStatus, 2000);
    }
}
