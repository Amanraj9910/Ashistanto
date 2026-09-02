import type { ActionPreview, ConfigStatus, TextMessageResponse, UserProfile, VoiceMessageResponse } from '@/types';
import { createId } from './utils';

export const mockConfig: ConfigStatus = {
  configured: true
};

export const mockUser: UserProfile = {
  displayName: 'User',
  firstName: 'User',
  email: 'user@company.onmicrosoft.com',
  role: 'Director of Ops'
};

function createActionPreview(text: string): ActionPreview {
  const lowerText = text.toLowerCase();

  if (lowerText.includes('email') || lowerText.includes('mail')) {
    return {
      actionId: createId('action'),
      title: 'Email Preview',
      type: 'email',
      status: 'pending',
      editable: ['recipientName', 'subject', 'body'],
      details: {
        to: 'Sarah Jenkins',
        subject: 'Project Follow Up',
        body: 'Hi Sarah, following up on the latest project update. Please share your current status when available.',
        cc: 'None'
      }
    };
  }

  if (lowerText.includes('meeting') || lowerText.includes('calendar') || lowerText.includes('schedule')) {
    return {
      actionId: createId('action'),
      title: 'Meeting Preview',
      type: 'meeting',
      status: 'pending',
      editable: ['subject', 'attendeeNames', 'startTime', 'endTime'],
      details: {
        subject: 'Budget Review',
        attendees: 'Sarah Jenkins, Design Team',
        startTime: 'Tomorrow, 11:00 AM',
        endTime: 'Tomorrow, 11:30 AM',
        isTeams: true
      }
    };
  }

  return {
    actionId: createId('action'),
    title: 'Leave Request Preview',
    type: 'leave',
    status: 'pending',
    editable: ['employee', 'leaveType', 'startDate'],
    details: {
      employee: 'Sarah Jenkins',
      leaveType: 'Sick Leave',
      startDate: 'Today',
      status: 'Ready for approval'
    }
  };
}

export async function mockSendTextMessage(text: string, sessionId: string): Promise<TextMessageResponse> {
  await new Promise((resolve) => setTimeout(resolve, 500));
  const lowerText = text.toLowerCase();

  if (
    lowerText.includes('send') ||
    lowerText.includes('schedule') ||
    lowerText.includes('leave') ||
    lowerText.includes('expense') ||
    lowerText.includes('mail')
  ) {
    return {
      success: true,
      sessionId,
      response: JSON.stringify({
        type: 'action_preview',
        preview: createActionPreview(text),
        message: 'I prepared a preview. Please review the details before I proceed.'
      })
    };
  }

  return {
    success: true,
    sessionId,
    response:
      'I reviewed the current workspace context. You can ask me to schedule meetings, draft messages, search enterprise files, or prepare operational workflows.'
  };
}

export async function mockProcessVoice(sessionId: string): Promise<VoiceMessageResponse> {
  await new Promise((resolve) => setTimeout(resolve, 650));

  return {
    transcript: 'Log a sick day for Sarah Jenkins starting today.',
    agentResponse: JSON.stringify({
      type: 'action_preview',
      preview: createActionPreview('leave request for Sarah Jenkins'),
      message: 'I prepared a leave request preview. Please confirm before I submit it.'
    }),
    sessionId
  };
}
