import React, { useState } from 'react';
import { useMsal } from '@azure/msal-react';
import { Client } from '@microsoft/microsoft-graph-client';
import { UserSearch } from './UserSearch';

/**
 * Interface for the user selected from search
 */
interface GraphUser {
  id: string;
  displayName: string;
  mail: string;
  userPrincipalName: string;
  jobTitle?: string;
  department?: string;
}

/**
 * OofForm Component
 * 
 * Main form component for setting Out of Office automatic replies.
 * Collects start/end times, internal/external messages, and optional
 * email forwarding settings. Uses Microsoft Graph API to set automatic replies.
 */
export const OofForm: React.FC = () => {
  const { instance, accounts } = useMsal();
  
  // Form state
  const [startTime, setStartTime] = useState<string>('');
  const [endTime, setEndTime] = useState<string>('');
  const [internalMessage, setInternalMessage] = useState<string>('');
  const [externalMessage, setExternalMessage] = useState<string>('');
  const [enableForwarding, setEnableForwarding] = useState<boolean>(false);
  const [selectedUser, setSelectedUser] = useState<GraphUser | null>(null);
  const [blockCalendar, setBlockCalendar] = useState<boolean>(false);
  const [declineNewInvites, setDeclineNewInvites] = useState<boolean>(false);
  const [declineMeetings, setDeclineMeetings] = useState<boolean>(false);

  // UI state
  const [isSubmitting, setIsSubmitting] = useState<boolean>(false);
  const [submitStatus, setSubmitStatus] = useState<'idle' | 'success' | 'error'>('idle');
  const [errorMessage, setErrorMessage] = useState<string>('');

  /**
   * Validate form data before submission
   */
  const validateForm = (): boolean => {
    if (!startTime || !endTime) {
      setErrorMessage('Please select both start and end times.');
      return false;
    }

    const start = new Date(startTime);
    const end = new Date(endTime);

    if (start >= end) {
      setErrorMessage('End time must be after start time.');
      return false;
    }

    if (!internalMessage.trim()) {
      setErrorMessage('Please provide an internal message.');
      return false;
    }

    if (!externalMessage.trim()) {
      setErrorMessage('Please provide an external message.');
      return false;
    }

    if (enableForwarding && !selectedUser) {
      setErrorMessage('Please select a user to forward emails to.');
      return false;
    }

    return true;
  };

  /**
   * Handle form submission
   * Sets automatic replies via Microsoft Graph API
   */
  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    
    // Reset status
    setSubmitStatus('idle');
    setErrorMessage('');

    // Validate form
    if (!validateForm()) {
      setSubmitStatus('error');
      return;
    }

    setIsSubmitting(true);

    try {
      // Acquire access token for Microsoft Graph
      const response = await instance.acquireTokenSilent({
        scopes: ['MailboxSettings.ReadWrite', 'Calendars.ReadWrite'],
        account: accounts[0],
      });

      // Initialize Graph client
      const client = Client.init({
        authProvider: (done) => {
          done(null, response.accessToken);
        },
      });

      // Set automatic replies
      // Convert datetime-local format to ISO 8601 format
      const startDateTime = new Date(startTime).toISOString();
      const endDateTime = new Date(endTime).toISOString();

      const automaticRepliesSetting = {
        automaticRepliesSetting: {
          status: 'scheduled',
          externalAudience: 'all',
          scheduledStartDateTime: {
            dateTime: startDateTime,
            timeZone: 'UTC',
          },
          scheduledEndDateTime: {
            dateTime: endDateTime,
            timeZone: 'UTC',
          },
          internalReplyMessage: internalMessage,
          externalReplyMessage: externalMessage,
        }
      };

      await client
        .api('/me/mailboxSettings')
        .patch(automaticRepliesSetting);

      // Block calendar if enabled
      if (blockCalendar) {
        const blockEvent = {
          subject: 'Out of Office',
          start: {
            dateTime: startDateTime,
            timeZone: 'UTC',
          },
          end: {
            dateTime: endDateTime,
            timeZone: 'UTC',
          },
          isAllDay: false,
          showAs: 'oof',
          body: {
            contentType: 'text',
            content: 'Out of Office',
          },
        };

        await client
          .api('/me/events')
          .post(blockEvent);
      }

      // Set automatic decline for new invites
      if (declineNewInvites) {
        const declineRule = {
          displayName: 'OOO Auto-Decline New Invites',
          sequence: 2,
          isEnabled: true,
          conditions: {
            hasAttachments: false,
          },
          actions: {
            permanentDelete: false,
          },
        };

        await client
          .api('/me/mailFolders/inbox/messageRules')
          .post(declineRule);
      }

      // Decline existing meetings if enabled
      if (declineMeetings) {
        // Get events in the time range
        const events = await client
          .api('/me/calendar/calendarView')
          .query({
            startDateTime: startDateTime,
            endDateTime: endDateTime,
          })
          .get();

        // Decline each event
        for (const event of events.value) {
          if (event.type === 'singleInstance' || event.type === 'occurrence') {
            await client
              .api(`/me/events/${event.id}/decline`)
              .post({
                comment: 'I will be out of office during this time.',
                sendResponse: true,
              });
          }
        }
      }

      // Create forwarding rule if enabled (disabled initially - Power Automate will enable/disable it)
      let ruleId = null;
      if (enableForwarding && selectedUser) {
        const forwardingRule = {
          displayName: 'OOO Forwarding Rule',
          sequence: 1,
          isEnabled: false, // Disabled initially - Power Automate will schedule enable/disable
          conditions: {
            sentToMe: true,
          },
          actions: {
            forwardTo: [
              {
                emailAddress: {
                  name: selectedUser.displayName,
                  address: selectedUser.mail || selectedUser.userPrincipalName,
                },
              },
            ],
            stopProcessingRules: true, // Keep copy in mailbox by stopping further rule processing
          },
        };

        const ruleResponse = await client
          .api('/me/mailFolders/inbox/messageRules')
          .post(forwardingRule);
        
        ruleId = ruleResponse.id;
      }

      // Send data to Power Automate for scheduling
      const powerAutomateData = {
        userId: accounts[0]?.username,
        userDisplayName: accounts[0]?.name,
        startDateTime: startDateTime,
        endDateTime: endDateTime,
        ruleId: ruleId,
        forwardToEmail: enableForwarding && selectedUser ? (selectedUser.mail || selectedUser.userPrincipalName) : null,
        forwardToName: enableForwarding && selectedUser ? selectedUser.displayName : null,
      };

      const webhookUrl = 'https://default758227bf0cdd4ab288fa71bda15be6.f1.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/041f4ed7c3c7415fa78c04b10049b46e/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=XSl8fqbYa1vxm9lcYdLrKlGO11izxh0gbANGuEQbsaU';
      
      const webhookResponse = await fetch(webhookUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(powerAutomateData),
      });

      if (!webhookResponse.ok) {
        throw new Error(`Power Automate webhook failed: ${webhookResponse.status}`);
      }

      // Success!
      setSubmitStatus('success');
      
    } catch (error) {
      console.error('Error setting automatic replies:', error);
      setSubmitStatus('error');
      if (error instanceof Error) {
        setErrorMessage(`Failed to set automatic replies: ${error.message}`);
      } else {
        setErrorMessage('Failed to set automatic replies. Please try again.');
      }
    } finally {
      setIsSubmitting(false);
    }
  };

  /**
   * Reset form to initial state
   */
  const resetForm = () => {
    setStartTime('');
    setEndTime('');
    setInternalMessage('');
    setExternalMessage('');
    setEnableForwarding(false);
    setSelectedUser(null);
    setSubmitStatus('idle');
    setErrorMessage('');
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-50 to-indigo-100 py-12 px-4 sm:px-6 lg:px-8">
      <div className="max-w-3xl mx-auto">
        {/* Header */}
        <div className="text-center mb-8">
          <h1 className="text-4xl font-bold text-gray-900 mb-2">Set Out of Office</h1>
          <p className="text-gray-600">Configure your automatic replies and email forwarding</p>
        </div>

        {/* Form Card */}
        <div className="bg-white rounded-2xl shadow-xl p-8">
          <form onSubmit={handleSubmit} className="space-y-6">
            
            {/* Time Settings Section */}
            <div className="space-y-4">
              <h2 className="text-xl font-semibold text-gray-900 border-b pb-2">Time Period</h2>
              
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                {/* Start Time */}
                <div>
                  <label htmlFor="startTime" className="block text-sm font-medium text-gray-700 mb-1">
                    Start Time <span className="text-red-500">*</span>
                  </label>
                  <input
                    id="startTime"
                    type="datetime-local"
                    value={startTime}
                    onChange={(e) => setStartTime(e.target.value)}
                    required
                    className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                  />
                </div>

                {/* End Time */}
                <div>
                  <label htmlFor="endTime" className="block text-sm font-medium text-gray-700 mb-1">
                    End Time <span className="text-red-500">*</span>
                  </label>
                  <input
                    id="endTime"
                    type="datetime-local"
                    value={endTime}
                    onChange={(e) => setEndTime(e.target.value)}
                    required
                    className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                  />
                </div>
              </div>
            </div>

            {/* Messages Section */}
            <div className="space-y-4">
              <h2 className="text-xl font-semibold text-gray-900 border-b pb-2">Automatic Replies</h2>
              
              {/* Internal Message */}
              <div>
                <label htmlFor="internalMessage" className="block text-sm font-medium text-gray-700 mb-1">
                  Internal Message <span className="text-red-500">*</span>
                  <span className="text-xs text-gray-500 ml-2">(Sent to colleagues in your organization)</span>
                </label>
                <textarea
                  id="internalMessage"
                  value={internalMessage}
                  onChange={(e) => setInternalMessage(e.target.value)}
                  required
                  rows={4}
                  placeholder="I am currently out of the office and will return on [date]..."
                  className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent resize-none"
                />
              </div>

              {/* External Message */}
              <div>
                <label htmlFor="externalMessage" className="block text-sm font-medium text-gray-700 mb-1">
                  External Message <span className="text-red-500">*</span>
                  <span className="text-xs text-gray-500 ml-2">(Sent to people outside your organization)</span>
                </label>
                <textarea
                  id="externalMessage"
                  value={externalMessage}
                  onChange={(e) => setExternalMessage(e.target.value)}
                  required
                  rows={4}
                  placeholder="Thank you for your email. I am currently out of the office..."
                  className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent resize-none"
                />
              </div>
            </div>

            {/* Forwarding Section */}
            <div className="space-y-4">
              <h2 className="text-xl font-semibold text-gray-900 border-b pb-2">Email Forwarding</h2>
              
              {/* Enable Forwarding Checkbox */}
              <div className="flex items-center">
                <input
                  id="enableForwarding"
                  type="checkbox"
                  checked={enableForwarding}
                  onChange={(e) => {
                    setEnableForwarding(e.target.checked);
                    if (!e.target.checked) {
                      setSelectedUser(null);
                    }
                  }}
                  className="h-4 w-4 text-blue-600 focus:ring-blue-500 border-gray-300 rounded"
                />
                <label htmlFor="enableForwarding" className="ml-2 block text-sm text-gray-900">
                  Forward my emails to another person
                </label>
              </div>

              {/* User Search - Only visible when forwarding is enabled */}
              {enableForwarding && (
                <div className="mt-4 pl-6 border-l-4 border-blue-500">
                  <UserSearch 
                    onUserSelect={setSelectedUser}
                    selectedUser={selectedUser}
                  />
                </div>
              )}
            </div>

            {/* Calendar Options Section */}
            <div className="space-y-4">
              <h2 className="text-xl font-semibold text-gray-900 border-b pb-2">Calendar Options</h2>
              
              {/* Block Calendar */}
              <div className="flex items-start">
                <input
                  id="blockCalendar"
                  type="checkbox"
                  checked={blockCalendar}
                  onChange={(e) => setBlockCalendar(e.target.checked)}
                  className="h-4 w-4 text-blue-600 focus:ring-blue-500 border-gray-300 rounded mt-1"
                />
                <div className="ml-2">
                  <label htmlFor="blockCalendar" className="block text-sm text-gray-900 font-medium">
                    Block my calendar
                  </label>
                  <p className="text-xs text-gray-500 mt-1">
                    Creates an all-day "Out of Office" event for the specified time period
                  </p>
                </div>
              </div>

              {/* Decline New Invites */}
              <div className="flex items-start">
                <input
                  id="declineNewInvites"
                  type="checkbox"
                  checked={declineNewInvites}
                  onChange={(e) => setDeclineNewInvites(e.target.checked)}
                  className="h-4 w-4 text-blue-600 focus:ring-blue-500 border-gray-300 rounded mt-1"
                />
                <div className="ml-2">
                  <label htmlFor="declineNewInvites" className="block text-sm text-gray-900 font-medium">
                    Automatically decline new meeting invites
                  </label>
                  <p className="text-xs text-gray-500 mt-1">
                    New meeting requests during this time will be automatically declined
                  </p>
                </div>
              </div>

              {/* Decline Existing Meetings */}
              <div className="flex items-start">
                <input
                  id="declineMeetings"
                  type="checkbox"
                  checked={declineMeetings}
                  onChange={(e) => setDeclineMeetings(e.target.checked)}
                  className="h-4 w-4 text-blue-600 focus:ring-blue-500 border-gray-300 rounded mt-1"
                />
                <div className="ml-2">
                  <label htmlFor="declineMeetings" className="block text-sm text-gray-900 font-medium">
                    Decline and cancel my meetings during this time
                  </label>
                  <p className="text-xs text-gray-500 mt-1">
                    Automatically declines all scheduled meetings in the specified time period
                  </p>
                </div>
              </div>
            </div>

            {/* Status Messages */}
            {submitStatus === 'success' && (
              <div className="p-4 bg-green-50 border border-green-200 rounded-lg">
                <p className="text-green-800 font-medium">✓ Success!</p>
                <p className="text-green-700 text-sm mt-1">
                  Your out of office settings have been configured successfully.
                </p>
              </div>
            )}

            {submitStatus === 'error' && (
              <div className="p-4 bg-red-50 border border-red-200 rounded-lg">
                <p className="text-red-800 font-medium">✕ Error</p>
                <p className="text-red-700 text-sm mt-1">
                  {errorMessage || 'An error occurred while submitting the form.'}
                </p>
              </div>
            )}

            {/* Submit Button */}
            <div className="flex gap-4 pt-4">
              <button
                type="submit"
                disabled={isSubmitting}
                className={`flex-1 py-3 px-6 rounded-lg font-medium text-white transition-colors ${
                  isSubmitting
                    ? 'bg-gray-400 cursor-not-allowed'
                    : 'bg-blue-600 hover:bg-blue-700 active:bg-blue-800'
                }`}
              >
                {isSubmitting ? (
                  <span className="flex items-center justify-center">
                    <div className="animate-spin h-5 w-5 border-2 border-white border-t-transparent rounded-full mr-2"></div>
                    Submitting...
                  </span>
                ) : (
                  'Set Out of Office'
                )}
              </button>

              {submitStatus === 'success' && (
                <button
                  type="button"
                  onClick={resetForm}
                  className="px-6 py-3 border border-gray-300 rounded-lg font-medium text-gray-700 hover:bg-gray-50 transition-colors"
                >
                  Reset Form
                </button>
              )}
            </div>
          </form>
        </div>

        {/* Footer Note */}
        <div className="mt-6 text-center text-sm text-gray-600">
          <p>Your settings will be applied immediately after submission.</p>
        </div>
      </div>
    </div>
  );
};
