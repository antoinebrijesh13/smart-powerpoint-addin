// PowerPoint Add-in Task Pane Logic

Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
        console.log('✓ Smart PowerPoint add-in loaded');
        initializeAddin();
    }
});

// Configuration
const SESSION_ID = 'demo_session';
const BACKEND_URL = 'ws://localhost:8000';
const STUDENT_URL = 'http://localhost:8080/student.html?session=' + SESSION_ID;

// WebSocket connection
let ws = null;
let currentSlideNumber = 1;
let totalSlides = 0;

// Initialize the add-in
function initializeAddin() {
    // Set session ID
    document.getElementById('sessionId').textContent = SESSION_ID;
    document.getElementById('studentLink').textContent = STUDENT_URL;
    
    // Copy button
    document.getElementById('copyBtn').addEventListener('click', copyStudentLink);
    
    // Connect to backend
    connectToBackend();
    
    // Get total slides
    getTotalSlides();
    
    // Monitor slide changes
    monitorSlideChanges();
    
    // Auto-detect slide on load
    detectCurrentSlide();
}

// Connect to WebSocket backend
function connectToBackend() {
    const wsUrl = `${BACKEND_URL}/ws/${SESSION_ID}`;
    
    updateStatus('Connecting...', false);
    
    try {
        ws = new WebSocket(wsUrl);
        
        ws.onopen = () => {
            console.log('✓ Connected to backend');
            updateStatus('Connected', true);
        };
        
        ws.onmessage = (event) => {
            const data = JSON.parse(event.data);
            handleBackendMessage(data);
        };
        
        ws.onerror = (error) => {
            console.error('WebSocket error:', error);
            updateStatus('Connection failed', false);
        };
        
        ws.onclose = () => {
            console.log('Disconnected from backend');
            updateStatus('Disconnected', false);
            
            // Attempt reconnection after 3 seconds
            setTimeout(connectToBackend, 3000);
        };
        
    } catch (error) {
        console.error('Failed to connect:', error);
        updateStatus('Error', false);
    }
}

// Handle messages from backend
function handleBackendMessage(data) {
    console.log('Backend message:', data);
    
    // Update student count
    if (data.student_count !== undefined) {
        document.getElementById('studentCount').textContent = data.student_count;
    }
    
    // Handle notification (current slide update)
    if (data.type === 'current_slide_update' && data.slide_summary) {
        displayNotification(data.slide_summary);
    }
}

// Monitor slide changes in PowerPoint
function monitorSlideChanges() {
    // Listen for selection change events (includes slide changes)
    Office.context.document.addHandlerAsync(
        Office.EventType.DocumentSelectionChanged,
        onSlideChange,
        (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                console.log('✓ Monitoring slide changes');
            } else {
                console.error('Failed to add slide change handler');
            }
        }
    );
}

// Called when slide changes
function onSlideChange(eventArgs) {
    detectCurrentSlide();
}

// Detect current slide number
function detectCurrentSlide() {
    Office.context.document.getSelectedDataAsync(
        Office.CoercionType.SlideRange,
        (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                if (result.value && result.value.slides && result.value.slides.length > 0) {
                    const slideIndex = result.value.slides[0].index;
                    currentSlideNumber = slideIndex + 1; // 0-indexed to 1-indexed
                    
                    // Update UI
                    document.getElementById('currentSlide').textContent = currentSlideNumber;
                    
                    // Notify backend of slide change
                    if (ws && ws.readyState === WebSocket.OPEN) {
                        ws.send(JSON.stringify({
                            type: 'slide_change',
                            slide_number: currentSlideNumber
                        }));
                        console.log(`Slide changed to ${currentSlideNumber}`);
                    }
                }
            } else {
                console.warn('Could not detect current slide');
            }
        }
    );
}

// Get total number of slides
function getTotalSlides() {
    // PowerPoint doesn't have a direct API for total slides
    // We can approximate by trying to select all slides
    Office.context.document.getFilePropertiesAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            // This is a workaround - actual slide count needs more complex logic
            // For now, display "--" and update when we detect higher slide numbers
            document.getElementById('totalSlides').textContent = '--';
        }
    });
}

// Update connection status
function updateStatus(text, isOnline) {
    document.getElementById('statusText').textContent = text;
    const dot = document.querySelector('.status-dot');
    
    if (isOnline) {
        dot.classList.remove('offline');
        dot.classList.add('online');
    } else {
        dot.classList.remove('online');
        dot.classList.add('offline');
    }
}

// Display notification
function displayNotification(summary) {
    const container = document.getElementById('notificationsContainer');
    
    // Remove "no notifications" message
    const noNotif = container.querySelector('.no-notifications');
    if (noNotif) {
        noNotif.remove();
    }
    
    // Create notification card
    const card = document.createElement('div');
    card.className = 'notification-card';
    
    // Determine priority (if provided)
    const priority = summary.priority || 'low';
    card.classList.add(priority);
    
    const time = new Date().toLocaleTimeString('en-US', { 
        hour: '2-digit', 
        minute: '2-digit'
    });
    
    card.innerHTML = `
        <div class="notification-header">
            <span class="notification-count">${summary.total_students || 0} students</span>
            <span class="notification-time">${time}</span>
        </div>
        <div class="notification-summary">
            ${summary.summary || 'Students have questions'}
        </div>
        ${summary.suggestion ? `
            <div class="notification-suggestion">
                💡 ${summary.suggestion}
            </div>
        ` : ''}
    `;
    
    // Insert at top
    container.insertBefore(card, container.firstChild);
    
    // Limit to 10 most recent notifications
    while (container.children.length > 10) {
        container.removeChild(container.lastChild);
    }
}

// Copy student link to clipboard
function copyStudentLink() {
    const link = document.getElementById('studentLink').textContent;
    
    // Use modern clipboard API
    if (navigator.clipboard) {
        navigator.clipboard.writeText(link).then(() => {
            showCopyFeedback();
        }).catch(err => {
            console.error('Failed to copy:', err);
            fallbackCopy(link);
        });
    } else {
        fallbackCopy(link);
    }
}

// Fallback copy method
function fallbackCopy(text) {
    const textarea = document.createElement('textarea');
    textarea.value = text;
    textarea.style.position = 'fixed';
    textarea.style.opacity = '0';
    document.body.appendChild(textarea);
    textarea.select();
    
    try {
        document.execCommand('copy');
        showCopyFeedback();
    } catch (err) {
        console.error('Fallback copy failed:', err);
    }
    
    document.body.removeChild(textarea);
}

// Show copy feedback
function showCopyFeedback() {
    const btn = document.getElementById('copyBtn');
    const originalText = btn.textContent;
    btn.textContent = '✓';
    btn.style.background = '#10b981';
    btn.style.color = 'white';
    
    setTimeout(() => {
        btn.textContent = originalText;
        btn.style.background = '';
        btn.style.color = '';
    }, 1500);
}

// Debug: Log PowerPoint info
console.log('Office.js version:', Office.context.diagnostics.version);
console.log('Host:', Office.context.diagnostics.host);
console.log('Platform:', Office.context.diagnostics.platform);
