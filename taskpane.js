/* global Office */

let salesforceSession = null;
let currentEmail = null;

// Initialize the add-in
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        console.log('Office initialized successfully');
        document.getElementById("connectionStatus").style.color = "#28a745";
        loadCurrentEmail();
        checkSalesforceAuth();
    }
});

// Check if user is already authenticated to Salesforce
function checkSalesforceAuth() {
    try {
        const savedSession = localStorage.getItem('salesforceSession');
        if (savedSession) {
            salesforceSession = JSON.parse(savedSession);
            if (salesforceSession.access_token && !isTokenExpired(salesforceSession)) {
                showMainApp();
                loadRelatedRecords();
                loadRecentActivity();
                return;
            }
        }
    } catch (e) {
        console.error('Error parsing saved session:', e);
    }
    showAuthSection();
}

// Check if access token is expired
function isTokenExpired(session) {
    if (!session.issued_at) return true;
    const issuedAt = parseInt(session.issued_at);
    const now = Date.now();
    const expirationTime = 2 * 60 * 60 * 1000; // 2 hours in milliseconds
    return (now - issuedAt) > expirationTime;
}

// Show authentication section
function showAuthSection() {
    document.getElementById('authSection').classList.remove('hidden');
    document.getElementById('mainApp').classList.add('hidden');
}

// Show main application
function showMainApp() {
    document.getElementById('authSection').classList.add('hidden');
    document.getElementById('mainApp').classList.remove('hidden');
    document.getElementById('connectionStatus').style.color = '#28a745';
}

// Authenticate to Salesforce using OAuth2
async function authenticateToSalesforce() {
    const instanceUrl = document.getElementById('instanceUrl').value;
    if (!instanceUrl) {
        alert('Please enter your Salesforce instance URL');
        return;
    }

    try {
        // For demo purposes, we'll simulate successful authentication
        // In a real implementation, you would implement proper OAuth2 flow
        await simulateAuthentication(instanceUrl);
        
    } catch (error) {
        console.error('Authentication failed:', error);
        alert('Authentication failed. Please check your credentials and try again.');
    }
}

// Simulate authentication process (replace with real OAuth2 flow)
async function simulateAuthentication(instanceUrl) {
    // This is a placeholder - in reality you'd implement proper OAuth2 flow
    // For now, we'll just show that the UI works
    const mockSession = {
        access_token: 'mock_access_token_' + Date.now(),
        instance_url: instanceUrl,
        issued_at: Date.now().toString(),
        signature: 'mock_signature',
        token_type: 'Bearer'
    };
    
    salesforceSession = mockSession;
    localStorage.setItem('salesforceSession', JSON.stringify(mockSession));
    
    showMainApp();
    loadRelatedRecords();
    loadRecentActivity();
    
    // Show success message
    alert('Connected to Salesforce successfully! (Demo mode - no real connection yet)');
}

// Load current email information
function loadCurrentEmail() {
    try {
        if (Office.context.mailbox.item.subject) {
            Office.context.mailbox.item.subject.getAsync((result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    document.getElementById('currentSubject').textContent = result.value || 'No Subject';
                    updateLogSubject(result.value);
                }
            });
        }
        
        if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
            const fromField = Office.context.mailbox.item.from;
            if (fromField) {
                document.getElementById('currentFrom').textContent = 
                    `From: ${fromField.displayName || fromField.emailAddress}`;
            }
        }
    } catch (error) {
        console.error('Error loading email info:', error);
        document.getElementById('currentSubject').textContent = 'Email information unavailable';
    }
}

// Update log subject field
function updateLogSubject(subject) {
    const logSubject = document.getElementById('logSubject');
    if (logSubject) {
        logSubject.value = subject || '';
    }
}

// Switch between tabs
function switchTab(tabName) {
    // Remove active class from all tabs
    document.querySelectorAll('.tab').forEach(tab => {
        tab.classList.remove('active');
    });
    
    // Hide all tab content
    document.querySelectorAll('.tab-content').forEach(content => {
        content.classList.add('hidden');
    });
    
    // Show selected tab
    document.querySelector(`[onclick="switchTab('${tabName}')"]`).classList.add('active');
    document.getElementById(`${tabName}Tab`).classList.remove('hidden');
}

// Log current email to Salesforce
async function logCurrentEmail() {
    if (!salesforceSession) {
        alert('Please authenticate to Salesforce first');
        return;
    }
    
    try {
        // Get email details
        const emailData = await getCurrentEmailData();
        
        // In a real implementation, you would make API call to Salesforce
        const result = await saveEmailToSalesforce(emailData);
        
        if (result.success) {
            alert('Email logged successfully to Salesforce (Demo mode)');
            loadRecentActivity(); // Refresh activity
        } else {
            alert('Failed to log email: ' + result.error);
        }
        
    } catch (error) {
        console.error('Error logging email:', error);
        alert('Error logging email to Salesforce');
    }
}

// Get current email data
async function getCurrentEmailData() {
    return new Promise((resolve) => {
        const emailData = {
            subject: '',
            from: '',
            to: '',
            body: '',
            date: new Date()
        };
        
        try {
            if (Office.context.mailbox.item.subject) {
                Office.context.mailbox.item.subject.getAsync((result) => {
                    emailData.subject = result.value || '';
                    
                    if (Office.context.mailbox.item.body) {
                        Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, (bodyResult) => {
                            emailData.body = bodyResult.value || '';
                            resolve(emailData);
                        });
                    } else {
                        resolve(emailData);
                    }
                });
            } else {
                resolve(emailData);
            }
        } catch (error) {
            console.error('Error getting email data:', error);
            resolve(emailData);
        }
    });
}

// Save email to Salesforce (mock implementation)
async function saveEmailToSalesforce(emailData) {
    // This would be a real Salesforce API call
    return new Promise((resolve) => {
        setTimeout(() => {
            resolve({
                success: true,
                id: 'mock_activity_id_' + Date.now()
            });
        }, 1000);
    });
}

// Search contacts in Salesforce
async function searchContacts(event) {
    if (event && event.key === 'Enter') {
        performSearch();
    }
}

async function performSearch() {
    const searchTerm = document.getElementById('searchInput').value.trim();
    if (!searchTerm) {
        alert('Please enter search terms');
        return;
    }
    
    if (!salesforceSession) {
        alert('Please authenticate to Salesforce first');
        return;
    }
    
    const resultsContainer = document.getElementById('searchResults');
    resultsContainer.innerHTML = '<div class="loading">Searching...</div>';
    
    try {
        const results = await searchSalesforceRecords(searchTerm);
        displaySearchResults(results);
    } catch (error) {
        console.error('Search error:', error);
        resultsContainer.innerHTML = '<div class="loading">Search failed. Please try again.</div>';
    }
}

// Search Salesforce records (mock implementation)
async function searchSalesforceRecords(searchTerm) {
    // This would be a real Salesforce SOSL/SOQL query
    return new Promise((resolve) => {
        setTimeout(() => {
            const mockResults = [
                {
                    type: 'Contact',
                    id: 'contact_1',
                    name: 'John Smith',
                    email: 'john.smith@example.com',
                    company: 'Acme Corp',
                    title: 'VP Sales'
                },
                {
                    type: 'Lead',
                    id: 'lead_1',
                    name: 'Jane Doe',
                    email: 'jane.doe@prospect.com',
                    company: 'Prospect Inc',
                    title: 'Marketing Director'
                }
            ].filter(record => 
                record.name.toLowerCase().includes(searchTerm.toLowerCase()) ||
                record.email.toLowerCase().includes(searchTerm.toLowerCase()) ||
                record.company.toLowerCase().includes(searchTerm.toLowerCase())
            );
            resolve(mockResults);
        }, 1000);
    });
}

// Display search results
function displaySearchResults(results) {
    const container = document.getElementById('searchResults');
    
    if (!results || results.length === 0) {
        container.innerHTML = '<div class="loading">No results found</div>';
        return;
    }
    
    let html = '';
    results.forEach(record => {
        html += `
            <div class="contact-item" onclick="selectRecord('${record.id}', '${record.type}')">
                <div class="contact-name">${record.name}</div>
                <div class="contact-details">
                    ${record.email} • ${record.company} • ${record.title}
                    <br><small>${record.type}</small>
                </div>
            </div>
        `;
    });
    
    container.innerHTML = html;
}

// Select a record
function selectRecord(recordId, recordType) {
    // Add to related records dropdown
    const relatedSelect = document.getElementById('relatedTo');
    const option = document.createElement('option');
    option.value = recordId;
    option.textContent = `${recordType}: ${recordId}`;
    relatedSelect.appendChild(option);
    relatedSelect.value = recordId;
    
    // Switch to log tab
    switchTab('log');
    
    alert(`Selected ${recordType} record: ${recordId}`);
}

// Load related records
async function loadRelatedRecords() {
    const container = document.getElementById('relatedRecords');
    
    if (!salesforceSession) {
        container.innerHTML = '<div class="loading">Please authenticate to Salesforce</div>';
        return;
    }
    
    try {
        // In real implementation, search for records related to current email
        const related = await findRelatedRecords();
        displayRelatedRecords(related);
    } catch (error) {
        console.error('Error loading related records:', error);
        container.innerHTML = '<div class="loading">Error loading related records</div>';
    }
}

// Find related records (mock implementation)
async function findRelatedRecords() {
    return new Promise((resolve) => {
        setTimeout(() => {
            resolve([
                { type: 'Contact', name: 'John Smith', id: 'contact_1' },
                { type: 'Opportunity', name: 'Q4 Deal', id: 'opp_1' }
            ]);
        }, 1000);
    });
}

// Display related records
function displayRelatedRecords(records) {
    const container = document.getElementById('relatedRecords');
    
    if (!records || records.length === 0) {
        container.innerHTML = '<div class="loading">No related records found</div>';
        return;
    }
    
    let html = '';
    records.forEach(record => {
        html += `
            <div class="contact-item" onclick="viewRecord('${record.id}')">
                <div class="contact-name">${record.name}</div>
                <div class="contact-details">${record.type}</div>
            </div>
        `;
    });
    
    container.innerHTML = html;
}

// View a record
function viewRecord(recordId) {
    // In a real implementation, this would open the record in Salesforce
    alert(`Opening record: ${recordId} (Demo mode)`);
}

// Load recent activity
async function loadRecentActivity() {
    const container = document.getElementById('recentActivity');
    
    if (!salesforceSession) {
        container.innerHTML = '<div class="loading">Please authenticate to Salesforce</div>';
        return;
    }
    
    try {
        const activities = await getRecentActivities();
        displayRecentActivity(activities);
    } catch (error) {
        console.error('Error loading recent activity:', error);
        container.innerHTML = '<div class="loading">Error loading recent activity</div>';
    }
}

// Get recent activities (mock implementation)
async function getRecentActivities() {
    return new Promise((resolve) => {
        setTimeout(() => {
            resolve([
                {
                    type: 'Email',
                    subject: 'Follow up on proposal',
                    date: new Date(Date.now() - 2 * 60 * 60 * 1000), // 2 hours ago
                    contact: 'John Smith'
                },
                {
                    type: 'Call',
                    subject: 'Discovery call with prospect',
                    date: new Date(Date.now() - 24 * 60 * 60 * 1000), // 1 day ago
                    contact: 'Jane Doe'
                },
                {
                    type: 'Meeting',
                    subject: 'Product demo',
                    date: new Date(Date.now() - 3 * 24 * 60 * 60 * 1000), // 3 days ago
                    contact: 'Mike Johnson'
                }
            ]);
        }, 1000);
    });
}

// Display recent activity
function displayRecentActivity(activities) {
    const container = document.getElementById('recentActivity');
    
    if (!activities || activities.length === 0) {
        container.innerHTML = '<div class="loading">No recent activity found</div>';
        return;
    }
    
    let html = '';
    activities.forEach(activity => {
        const timeAgo = getTimeAgo(activity.date);
        html += `
            <div class="activity-item">
                <div class="activity-type">${activity.type}</div>
                <div>${activity.subject}</div>
                <div class="activity-date">${timeAgo} • ${activity.contact}</div>
            </div>
        `;
    });
    
    container.innerHTML = html;
}

// Get time ago string
function getTimeAgo(date) {
    const now = new Date();
    const diffMs = now - date;
    const diffHours = Math.floor(diffMs / (1000 * 60 * 60));
    const diffDays = Math.floor(diffHours / 24);
    
    if (diffDays > 0) {
        return `${diffDays} day${diffDays > 1 ? 's' : ''} ago`;
    } else if (diffHours > 0) {
        return `${diffHours} hour${diffHours > 1 ? 's' : ''} ago`;
    } else {
        const diffMins = Math.floor(diffMs / (1000 * 60));
        return `${diffMins} minute${diffMins > 1 ? 's' : ''} ago`;
    }
}

// Create contact
function createContact() {
    if (!salesforceSession) {
        alert('Please authenticate to Salesforce first');
        return;
    }
    
    // In demo mode, just show what would happen
    alert('Would open Salesforce to create a new Contact (Demo mode)');
}

// Create lead
function createLead() {
    if (!salesforceSession) {
        alert('Please authenticate to Salesforce first');
        return;
    }
    
    // In demo mode, just show what would happen
    alert('Would open Salesforce to create a new Lead (Demo mode)');
}

// Save activity
async function saveActivity() {
    const subject = document.getElementById('logSubject').value;
    const relatedTo = document.getElementById('relatedTo').value;
    const comments = document.getElementById('logComments').value;
    
    if (!subject.trim()) {
        alert('Please enter a subject');
        return;
    }
    
    if (!salesforceSession) {
        alert('Please authenticate to Salesforce first');
        return;
    }
    
    try {
        const emailData = await getCurrentEmailData();
        const activityData = {
            subject: subject,
            relatedTo: relatedTo,
            comments: comments,
            emailData: emailData
        };
        
        const result = await saveActivityToSalesforce(activityData);
        
        if (result.success) {
            alert('Activity saved successfully (Demo mode)');
            // Clear form
            document.getElementById('logComments').value = '';
            // Refresh activity list
            loadRecentActivity();
            // Switch back to overview
            switchTab('overview');
        } else {
            alert('Failed to save activity: ' + result.error);
        }
        
    } catch (error) {
        console.error('Error saving activity:', error);
        alert('Error saving activity to Salesforce');
    }
}

// Save activity to Salesforce (mock implementation)
async function saveActivityToSalesforce(activityData) {
    // This would be a real Salesforce API call to create an Activity/Task record
    return new Promise((resolve) => {
        setTimeout(() => {
            resolve({
                success: true,
                id: 'activity_' + Date.now()
            });
        }, 1000);
    });
}

// Logout from Salesforce
function logoutFromSalesforce() {
    localStorage.removeItem('salesforceSession');
    salesforceSession = null;
    showAuthSection();
}