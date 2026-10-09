document.addEventListener('DOMContentLoaded', () => {
    const chatInput = document.getElementById('chat-input');
    const sendBtn = document.getElementById('send-btn');
    const chatContainer = document.getElementById('chat-container');
    const welcomeScreen = document.querySelector('.welcome-screen');
    const inputFilesList = document.getElementById('input-files-list');
    const outputFilesList = document.getElementById('output-files-list');
    const clearChatBtn = document.getElementById('clear-chat-btn');
    const dashboardPanel = document.getElementById('dashboard-panel');
    const dashboardContent = document.getElementById('dashboard-content');
    const closeDashboardBtn = document.getElementById('close-dashboard');

    let isWaiting = false;

    // Initialize Markdown parser options
    marked.setOptions({
        breaks: true,
        gfm: true
    });

    // Auto-resize textarea
    chatInput.addEventListener('input', function() {
        this.style.height = 'auto';
        this.style.height = (this.scrollHeight) + 'px';
        if (this.value.trim() !== '') {
            sendBtn.disabled = false;
        } else {
            sendBtn.disabled = true;
        }
    });

    // Handle Enter key (Shift+Enter for new line)
    chatInput.addEventListener('keydown', function(e) {
        if (e.key === 'Enter' && !e.shiftKey) {
            e.preventDefault();
            sendMessage();
        }
    });

    sendBtn.addEventListener('click', sendMessage);

    clearChatBtn.addEventListener('click', async () => {
        if (confirm("Clear conversation history?")) {
            await fetch('/api/chat', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ message: "clear" })
            });
            
            // Remove all messages except welcome screen
            const messages = chatContainer.querySelectorAll('.message');
            messages.forEach(msg => msg.remove());
            welcomeScreen.style.display = 'block';
            dashboardPanel.classList.remove('open');
            dashboardContent.innerHTML = '<p class="empty-state">No visualizations generated yet.</p>';
        }
    });

    closeDashboardBtn.addEventListener('click', () => {
        dashboardPanel.classList.remove('open');
    });

    async function fetchFiles() {
        try {
            const res = await fetch('/api/files');
            const data = await res.json();
            
            updateFileList(inputFilesList, data.input_files, 'No input files');
            updateFileList(outputFilesList, data.output_files, 'No output files');
            
            // Check for images in output files to show in dashboard
            const images = data.output_files.filter(f => f.match(/\.(png|jpg|jpeg|gif)$/i));
            if (images.length > 0) {
                updateDashboard(images);
            }
        } catch (error) {
            console.error("Error fetching files:", error);
        }
    }

    function updateFileList(ulElement, files, emptyText) {
        ulElement.innerHTML = '';
        if (!files || files.length === 0) {
            ulElement.innerHTML = `<li style="cursor:default; opacity:0.5;">${emptyText}</li>`;
            return;
        }
        
        files.forEach(file => {
            const li = document.createElement('li');
            li.textContent = file;
            if (file.match(/\.(png|jpg|jpeg|gif)$/i)) {
                li.addEventListener('click', () => {
                    updateDashboard([file]);
                    dashboardPanel.classList.add('open');
                });
            }
            ulElement.appendChild(li);
        });
    }

    function updateDashboard(images) {
        if (dashboardContent.querySelector('.empty-state')) {
            dashboardContent.innerHTML = '';
        }
        
        // Add new images that aren't already there
        images.forEach(img => {
            const imgPath = `/output/${img}`;
            if (!dashboardContent.querySelector(`img[src="${imgPath}"]`)) {
                const card = document.createElement('div');
                card.className = 'vis-card';
                card.innerHTML = `
                    <img src="${imgPath}" alt="${img}" onclick="window.open('${imgPath}', '_blank')">
                    <p>${img}</p>
                `;
                dashboardContent.prepend(card); // Add to top
                dashboardPanel.classList.add('open');
            }
        });
    }

    // Expose to window for suggestion buttons
    window.setInputValue = function(val) {
        chatInput.value = val;
        chatInput.style.height = 'auto';
        sendBtn.disabled = false;
        chatInput.focus();
    };

    function addMessage(content, sender = 'user') {
        if (welcomeScreen.style.display !== 'none') {
            welcomeScreen.style.display = 'none';
        }

        const msgDiv = document.createElement('div');
        msgDiv.className = `message ${sender}`;
        
        const contentDiv = document.createElement('div');
        contentDiv.className = 'message-content';
        
        if (sender === 'bot') {
            contentDiv.innerHTML = marked.parse(content);
        } else {
            contentDiv.textContent = content;
        }
        
        msgDiv.appendChild(contentDiv);
        chatContainer.appendChild(msgDiv);
        scrollToBottom();
    }

    function showTyping() {
        const div = document.createElement('div');
        div.className = 'typing-indicator';
        div.id = 'typing-indicator';
        div.innerHTML = `
            <div class="typing-dot"></div>
            <div class="typing-dot"></div>
            <div class="typing-dot"></div>
        `;
        chatContainer.appendChild(div);
        scrollToBottom();
    }

    function hideTyping() {
        const indicator = document.getElementById('typing-indicator');
        if (indicator) {
            indicator.remove();
        }
    }

    function scrollToBottom() {
        chatContainer.scrollTop = chatContainer.scrollHeight;
    }

    async function sendMessage() {
        if (isWaiting) return;
        
        const text = chatInput.value.trim();
        if (!text) return;
        
        // UI Updates
        chatInput.value = '';
        chatInput.style.height = 'auto';
        sendBtn.disabled = true;
        addMessage(text, 'user');
        showTyping();
        isWaiting = true;
        
        try {
            const response = await fetch('/api/chat', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({ message: text })
            });
            
            const data = await response.json();
            hideTyping();
            
            if (response.ok) {
                addMessage(data.response, 'bot');
                // Refresh files after a message since the bot might have created something
                fetchFiles();
            } else {
                addMessage(`**Error:** ${data.error || 'Something went wrong.'}`, 'bot');
            }
        } catch (error) {
            hideTyping();
            addMessage(`**Network Error:** Could not connect to the server.`, 'bot');
            console.error(error);
        } finally {
            isWaiting = false;
            chatInput.focus();
        }
    }

    // Initial fetch
    fetchFiles();
    
    // Poll for files every 5 seconds just in case
    setInterval(fetchFiles, 5000);
});
