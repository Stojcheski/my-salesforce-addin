/**
 * Salesforce API Service
 * Handles all interactions with Salesforce REST API
 */

class SalesforceService {
    constructor() {
        this.baseUrl = null;
        this.accessToken = null;
        this.instanceUrl = null;
        this.apiVersion = 'v58.0';
    }

    /**
     * Initialize OAuth2 authentication flow
     * @param {string} instanceUrl - Salesforce instance URL
     * @param {string} clientId - Connected App Client ID
     * @param {string} redirectUri - OAuth redirect URI
     */
    async authenticate(instanceUrl, clientId, redirectUri) {
        this.instanceUrl = instanceUrl;
        
        // Construct OAuth URL
        const oauthUrl = `${instanceUrl}/services/oauth2/authorize?` +
            `response_type=code&` +
            `client_id=${encodeURIComponent(clientId)}&` +
            `redirect_uri=${encodeURIComponent(redirectUri)}&` +
            `scope=full refresh_token`;

        // In a real implementation using a popup or redirect
        return new Promise((resolve, reject) => {
            // For demo purposes, we'll use a popup window
            const popup = window.open(oauthUrl, 'salesforce-auth', 
                'width=600,height=600,scrollbars=yes,resizable=yes');
            
            // Monitor popup for OAuth callback
            const checkClosed = setInterval(() => {
                if (popup.closed) {
                    clearInterval(checkClosed);
                    reject(new Error('Authentication cancelled'));
                }
                
                try {
                    if (popup.location.href.includes(redirectUri)) {
                        const url = new URL(popup.location.href);
                        const code = url.searchParams.get('code');
                        const error = url.searchParams.get('error');
                        
                        popup.close();
                        clearInterval(checkClosed);
                        
                        if (error) {
                            reject(new Error(`OAuth error: ${error}`));
                        } else if (code) {
                            // Exchange code for access token
                            this.exchangeCodeForToken(code, clientId, redirectUri)
                                .then(resolve)
                                .catch(reject);
                        }
                    }
                } catch (e) {
                    // Cross-origin error - popup still on Salesforce domain
                }
            }, 1000);
        });
    }

    /**
     * Exchange authorization code for access token
     */
    async exchangeCodeForToken(code, clientId, redirectUri) {
        const tokenUrl = `${this.instanceUrl}/services/oauth2/token`;
        
        const params = new URLSearchParams({
            grant_type: 'authorization_code',
            code: code,
            client_id: clientId,
            redirect_uri: redirectUri
        });

        const response = await fetch(tokenUrl, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded'
            },
            body: params
        });

        if (!response.ok) {
            throw new Error(`Token exchange failed: ${response.status}`);
        }

        const tokenData = await response.json();
        
        this.accessToken = tokenData.access_token;
        this.instanceUrl = tokenData.instance_url;
        
        // Store in localStorage for persistence
        const sessionData = {
            access_token: tokenData.access_token,
            refresh_token: tokenData.refresh_token,
            instance_url: tokenData.instance_url,
            issued_at: tokenData.issued_at,
            signature: tokenData.signature
        };
        
        localStorage.setItem('salesforceSession', JSON.stringify(sessionData));
        
        return sessionData;
    }

    /**
     * Refresh access token using refresh token
     */
    async refreshToken(clientId) {
        const session = JSON.parse(localStorage.getItem('salesforceSession'));
        if (!session || !session.refresh_token) {
            throw new Error('No refresh token available');
        }

        const tokenUrl = `${this.instanceUrl}/services/oauth2/token`;
        
        const params = new URLSearchParams({
            grant_type: 'refresh_token',
            refresh_token: session.refresh_token,
            client_id: clientId
        });

        const response = await fetch(tokenUrl, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded'
            },
            body: params
        });

        if (!response.ok) {
            // Refresh token expired, need full re-authentication
            localStorage.removeItem('salesforceSession');
            throw new Error('Refresh token expired');
        }

        const tokenData = await response.json();
        this.accessToken = tokenData.access_token;
        
        // Update stored session
        session.access_token = tokenData.access_token;
        session.issued_at = tokenData.issued_at;
        localStorage.setItem('salesforceSession', JSON.stringify(session));
        
        return tokenData;
    }

    /**
     * Make authenticated API call to Salesforce
     */
    async apiCall(endpoint, method = 'GET', data = null) {
        if (!this.accessToken) {
            const session = JSON.parse(localStorage.getItem('salesforceSession'));
            if (session && session.access_token) {
                this.accessToken = session.access_token;
                this.instanceUrl = session.instance_url;
            } else {
                throw new Error('Not authenticated');
            }
        }

        const url = `${this.instanceUrl}/services/data/${this.apiVersion}/${endpoint}`;
        
        const options = {
            method: method,
            headers: {
                'Authorization': `Bearer ${this.accessToken}`,
                'Content-Type': 'application/json'
            }
        };

        if (data && (method === 'POST' || method === 'PATCH' || method === 'PUT')) {
            options.body = JSON.stringify(data);
        }

        const response = await fetch(url, options);
        
        if (response.status === 401) {
            // Token expired, try to refresh
            try {
                await this.refreshToken();
                // Retry the original request
                options.headers['Authorization'] = `Bearer ${this.accessToken}`;
                const retryResponse = await fetch(url, options);
                if (!retryResponse.ok) {
                    throw new Error(`API call failed: ${retryResponse.status}`);
                }
                return await retryResponse.json();
            } catch (refreshError) {
                throw new Error('Authentication expired. Please re-authenticate.');
            }
        }

        if (!response.ok) {
            const errorText = await response.text();
            throw new Error(`API call failed: ${response.status} - ${errorText}`);
        }

        // Handle empty responses (like DELETE operations)
        if (response.status === 204) {
            return {};
        }

        return await response.json();
    }

    /**
     * Search for records using SOSL
     */
    async search(searchTerm, objectTypes = ['Contact', 'Lead', 'Account']) {
        const sosl = `FIND {${searchTerm}} IN ALL FIELDS RETURNING ${objectTypes.map(type => 
            `${type}(Id, Name, Email)`
        ).join(', ')}`;
        
        const encodedSOSL = encodeURIComponent(sosl);
        return await this.apiCall(`search/?q=${encodedSOSL}`);
    }

    /**
     * Query records using SOQL
     */
    async query(soql) {
        const encodedSOQL = encodeURIComponent(soql);
        return await this.apiCall(`query/?q=${encodedSOQL}`);
    }

    /**
     * Create a new record
     */
    async createRecord(objectType, recordData) {
        return await this.apiCall(`sobjects/${objectType}/`, 'POST', recordData);
    }

    /**
     * Update an existing record
     */
    async updateRecord(objectType, recordId, recordData) {
        return await this.apiCall(`sobjects/${objectType}/${recordId}`, 'PATCH', recordData);
    }

    /**
     * Get a record by ID
     */
    async getRecord(objectType, recordId, fields = null) {
        let endpoint = `sobjects/${objectType}/${recordId}`;
        if (fields) {
            endpoint += `?fields=${fields.join(',')}`;
        }
        return await this.apiCall(endpoint);
    }

    /**
     * Delete a record
     */
    async deleteRecord(objectType, recordId) {
        return await this.apiCall(`sobjects/${objectType}/${recordId}`, 'DELETE');
    }

    /**
     * Log email as an EmailMessage record
     */
    async logEmail(emailData, relatedRecordId = null) {
        const emailRecord = {
            Subject: emailData.subject,
            TextBody: emailData.body,
            FromAddress: emailData.from,
            ToAddress: emailData.to,
            MessageDate: emailData.date,
            Status: '3', // Sent
            Incoming: emailData.incoming || false
        };

        if (relatedRecordId) {
            emailRecord.RelatedToId = relatedRecordId;
        }

        return await this.createRecord('EmailMessage', emailRecord);
    }

    /**
     * Create a Task record for email activity
     */
    async createEmailTask(emailData, relatedRecordId = null, contactId = null) {
        const taskRecord = {
            Subject: `Email: ${emailData.subject}`,
            Description: emailData.body,
            ActivityDate: new Date().toISOString().split('T')[0], // Today's date
            Status: 'Completed',
            Type: 'Email',
            Priority: 'Normal'
        };

        if (relatedRecordId) {
            taskRecord.WhatId = relatedRecordId; // Related to Account, Opportunity, etc.
        }

        if (contactId) {
            taskRecord.WhoId = contactId; // Related to Contact or Lead
        }

        return await this.createRecord('Task', taskRecord);
    }

    /**
     * Find contacts or leads by email address
     */
    async findByEmail(emailAddress) {
        const results = {
            contacts: [],
            leads: []
        };

        try {
            // Search Contacts
            const contactQuery = `SELECT Id, Name, Email, Account.Name, Title FROM Contact WHERE Email = '${emailAddress}'`;
            const contactResult = await this.query(contactQuery);
            results.contacts = contactResult.records || [];

            // Search Leads
            const leadQuery = `SELECT Id, Name, Email, Company, Title FROM Lead WHERE Email = '${emailAddress}' AND IsConverted = false`;
            const leadResult = await this.query(leadQuery);
            results.leads = leadResult.records || [];

        } catch (error) {
            console.error('Error finding records by email:', error);
        }

        return results;
    }

    /**
     * Get recent activities
     */
    async getRecentActivities(limit = 10) {
        const userId = await this.getCurrentUserId();
        const query = `SELECT Id, Subject, ActivityDate, Type, Status, Who.Name, What.Name 
                      FROM Task 
                      WHERE OwnerId = '${userId}' 
                      ORDER BY CreatedDate DESC 
                      LIMIT ${limit}`;
        
        return await this.query(query);
    }

    /**
     * Get current user ID
     */
    async getCurrentUserId() {
        const userInfo = await this.apiCall('sobjects/User/me');
        return userInfo.Id;
    }

    /**
     * Get organization information
     */
    async getOrgInfo() {
        return await this.apiCall('sobjects/Organization');
    }

    /**
     * Search contacts and leads by various criteria
     */
    async searchContactsAndLeads(searchTerm) {
        const results = [];
        
        try {
            // Search contacts
            const contactQuery = `SELECT Id, Name, Email, Phone, Account.Name, Title 
                                 FROM Contact 
                                 WHERE Name LIKE '%${searchTerm}%' 
                                 OR Email LIKE '%${searchTerm}%' 
                                 OR Account.Name LIKE '%${searchTerm}%'
                                 LIMIT 10`;
            
            const contactResult = await this.query(contactQuery);
            if (contactResult.records) {
                contactResult.records.forEach(contact => {
                    results.push({
                        id: contact.Id,
                        name: contact.Name,
                        email: contact.Email,
                        phone: contact.Phone,
                        company: contact.Account ? contact.Account.Name : '',
                        title: contact.Title,
                        type: 'Contact'
                    });
                });
            }

            // Search leads
            const leadQuery = `SELECT Id, Name, Email, Phone, Company, Title 
                              FROM Lead 
                              WHERE IsConverted = false 
                              AND (Name LIKE '%${searchTerm}%' 
                              OR Email LIKE '%${searchTerm}%' 
                              OR Company LIKE '%${searchTerm}%')
                              LIMIT 10`;
            
            const leadResult = await this.query(leadQuery);
            if (leadResult.records) {
                leadResult.records.forEach(lead => {
                    results.push({
                        id: lead.Id,
                        name: lead.Name,
                        email: lead.Email,
                        phone: lead.Phone,
                        company: lead.Company,
                        title: lead.Title,
                        type: 'Lead'
                    });
                });
            }

        } catch (error) {
            console.error('Error searching contacts and leads:', error);
        }

        return results;
    }

    /**
     * Create a new Contact record
     */
    async createContact(contactData) {
        const contact = {
            FirstName: contactData.firstName,
            LastName: contactData.lastName,
            Email: contactData.email,
            Phone: contactData.phone,
            Title: contactData.title,
            Department: contactData.department,
            Description: contactData.description
        };

        // Remove empty fields
        Object.keys(contact).forEach(key => {
            if (!contact[key]) {
                delete contact[key];
            }
        });

        return await this.createRecord('Contact', contact);
    }

    /**
     * Create a new Lead record
     */
    async createLead(leadData) {
        const lead = {
            FirstName: leadData.firstName,
            LastName: leadData.lastName,
            Email: leadData.email,
            Phone: leadData.phone,
            Company: leadData.company,
            Title: leadData.title,
            Status: leadData.status || 'Open - Not Contacted',
            LeadSource: leadData.source || 'Email',
            Description: leadData.description
        };

        // Remove empty fields
        Object.keys(lead).forEach(key => {
            if (!lead[key]) {
                delete lead[key];
            }
        });

        return await this.createRecord('Lead', lead);
    }

    /**
     * Get account information by ID
     */
    async getAccount(accountId) {
        const fields = ['Id', 'Name', 'Type', 'Industry', 'Phone', 'Website', 'BillingCity', 'BillingState'];
        return await this.getRecord('Account', accountId, fields);
    }

    /**
     * Get opportunity information by ID
     */
    async getOpportunity(opportunityId) {
        const fields = ['Id', 'Name', 'StageName', 'Amount', 'CloseDate', 'Account.Name', 'Owner.Name'];
        return await this.getRecord('Opportunity', opportunityId, fields);
    }

    /**
     * Search for related records based on email participants
     */
    async findRelatedRecords(emailAddresses) {
        const results = {
            contacts: [],
            leads: [],
            accounts: [],
            opportunities: []
        };

        if (!emailAddresses || emailAddresses.length === 0) {
            return results;
        }

        try {
            const emailList = emailAddresses.map(email => `'${email}'`).join(',');

            // Find contacts
            const contactQuery = `SELECT Id, Name, Email, Account.Name, Account.Id, Title 
                                 FROM Contact 
                                 WHERE Email IN (${emailList})`;
            
            const contactResult = await this.query(contactQuery);
            if (contactResult.records) {
                results.contacts = contactResult.records;
                
                // Get related accounts from contacts
                const accountIds = contactResult.records
                    .filter(contact => contact.Account && contact.Account.Id)
                    .map(contact => contact.Account.Id);
                
                if (accountIds.length > 0) {
                    const uniqueAccountIds = [...new Set(accountIds)];
                    const accountList = uniqueAccountIds.map(id => `'${id}'`).join(',');
                    
                    const accountQuery = `SELECT Id, Name, Type, Industry 
                                         FROM Account 
                                         WHERE Id IN (${accountList})`;
                    
                    const accountResult = await this.query(accountQuery);
                    if (accountResult.records) {
                        results.accounts = accountResult.records;
                    }

                    // Get related opportunities
                    const oppQuery = `SELECT Id, Name, StageName, Amount, CloseDate, Account.Name 
                                     FROM Opportunity 
                                     WHERE AccountId IN (${accountList}) 
                                     AND IsClosed = false 
                                     ORDER BY CloseDate ASC 
                                     LIMIT 10`;
                    
                    const oppResult = await this.query(oppQuery);
                    if (oppResult.records) {
                        results.opportunities = oppResult.records;
                    }
                }
            }

            // Find leads (not converted)
            const leadQuery = `SELECT Id, Name, Email, Company, Title, Status 
                              FROM Lead 
                              WHERE Email IN (${emailList}) 
                              AND IsConverted = false`;
            
            const leadResult = await this.query(leadQuery);
            if (leadResult.records) {
                results.leads = leadResult.records;
            }

        } catch (error) {
            console.error('Error finding related records:', error);
        }

        return results;
    }

    /**
     * Get user's recent activities and tasks
     */
    async getUserActivities(limit = 20) {
        try {
            const userId = await this.getCurrentUserId();
            
            const query = `SELECT Id, Subject, Description, ActivityDate, Status, Type, Priority,
                          Who.Name, Who.Type, What.Name, What.Type, CreatedDate
                          FROM Task 
                          WHERE OwnerId = '${userId}'
                          ORDER BY CreatedDate DESC 
                          LIMIT ${limit}`;
            
            return await this.query(query);
        } catch (error) {
            console.error('Error getting user activities:', error);
            return { records: [] };
        }
    }

    /**
     * Create activity/task with email context
     */
    async createActivityFromEmail(emailData, activityData) {
        const task = {
            Subject: activityData.subject || `Email: ${emailData.subject}`,
            Description: activityData.description || emailData.body,
            ActivityDate: new Date().toISOString().split('T')[0],
            Status: activityData.status || 'Completed',
            Type: activityData.type || 'Email',
            Priority: activityData.priority || 'Normal'
        };

        // Link to related record if provided
        if (activityData.relatedToId) {
            // Determine if it's a Who (Contact/Lead) or What (Account/Opportunity) relationship
            if (activityData.relatedToType === 'Contact' || activityData.relatedToType === 'Lead') {
                task.WhoId = activityData.relatedToId;
            } else {
                task.WhatId = activityData.relatedToId;
            }
        }

        return await this.createRecord('Task', task);
    }

    /**
     * Validate and test connection
     */
    async testConnection() {
        try {
            const userInfo = await this.apiCall('sobjects/User/me');
            const orgInfo = await this.getOrgInfo();
            
            return {
                success: true,
                user: {
                    id: userInfo.Id,
                    name: userInfo.Name,
                    email: userInfo.Email,
                    username: userInfo.Username
                },
                org: {
                    name: orgInfo.records && orgInfo.records.length > 0 ? orgInfo.records[0].Name : 'Unknown',
                    id: orgInfo.records && orgInfo.records.length > 0 ? orgInfo.records[0].Id : 'Unknown'
                }
            };
        } catch (error) {
            return {
                success: false,
                error: error.message
            };
        }
    }

    /**
     * Logout and clear session
     */
    logout() {
        this.accessToken = null;
        this.instanceUrl = null;
        localStorage.removeItem('salesforceSession');
    }

    /**
     * Check if user is authenticated
     */
    isAuthenticated() {
        const session = localStorage.getItem('salesforceSession');
        if (!session) return false;
        
        try {
            const sessionData = JSON.parse(session);
            return !!(sessionData.access_token && sessionData.instance_url);
        } catch (e) {
            return false;
        }
    }

    /**
     * Get current session info
     */
    getSessionInfo() {
        const session = localStorage.getItem('salesforceSession');
        if (!session) return null;
        
        try {
            return JSON.parse(session);
        } catch (e) {
            return null;
        }
    }
}

if (typeof module !== 'undefined' && module.exports) {
    module.exports = SalesforceService;
} else {
    window.SalesforceService = SalesforceService;
}