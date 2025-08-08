import { meet_v2 } from "@googleapis/meet";

import { CalendarAuth } from "@calcom/app-store/googlecalendar/lib/CalendarAuth";
import logger from "@calcom/lib/logger";
import { getPiiFreeCalendarEvent } from "@calcom/lib/piiFreeData";
import { safeStringify } from "@calcom/lib/safeStringify";
import prisma from "@calcom/prisma";
import type { CalendarEvent } from "@calcom/types/Calendar";
import type { CredentialPayload } from "@calcom/types/Credential";
import type { CredentialForCalendarServiceWithEmail } from "@calcom/types/Credential";
import type { PartialReference } from "@calcom/types/EventManager";
import type { VideoApiAdapter, VideoCallData } from "@calcom/types/VideoApiAdapter";

const log = logger.getSubLogger({ prefix: ["app-store/googlevideo/lib/VideoApiAdapter"] });

const GoogleMeetVideoApiAdapter = (credential: CredentialPayload): VideoApiAdapter => {
  const getMeetClient = async () => {
    const googleCalendarCredential = await prisma.credential.findFirst({
      where: { userId: credential.userId ?? undefined, type: "google_calendar" },
      include: { user: true, delegatedTo: true },
    });
    if (!googleCalendarCredential) {
      throw new Error("Google Calendar credential not found");
    }
    const calendarAuth = new CalendarAuth(googleCalendarCredential as CredentialForCalendarServiceWithEmail);
    const calendarClient = await calendarAuth.getClient();
    const authClient = calendarClient.context._options.auth as any;
    return new meet_v2.Meet({ auth: authClient });
  };

  return {
    getAvailability: async () => [],
    createMeeting: async (event: CalendarEvent): Promise<VideoCallData> => {
      try {
        const meet = await getMeetClient();
        const { data } = await meet.spaces.create({});
        if (!data.meetingUri || !data.name) {
          throw new Error("Invalid response from Google Meet");
        }
        return {
          type: "google_video",
          id: data.name,
          url: data.meetingUri,
          password: "",
        };
      } catch (err) {
        log.error(
          "Google Meet meeting creation failed",
          safeStringify({ error: err, event: getPiiFreeCalendarEvent(event) })
        );
        throw new Error("Unexpected error");
      }
    },
    deleteMeeting: async (uid: string): Promise<void> => {
      try {
        const meet = await getMeetClient();
        await meet.spaces.endActiveConference({ name: uid });
      } catch (err) {
        log.error("Failed to delete meeting", safeStringify(err));
        return Promise.reject(new Error("Failed to delete meeting"));
      }
    },
    updateMeeting: async (bookingRef: PartialReference, event: CalendarEvent): Promise<VideoCallData> => {
      try {
        const meet = await getMeetClient();
        const { data } = await meet.spaces.get({ name: bookingRef.uid });
        if (!data.meetingUri || !data.name) {
          throw new Error("Invalid response from Google Meet");
        }
        return {
          type: "google_video",
          id: data.name,
          url: data.meetingUri,
          password: "",
        };
      } catch (err) {
        log.error(
          "Failed to update meeting",
          safeStringify({ error: err, event: getPiiFreeCalendarEvent(event) })
        );
        return Promise.reject(new Error("Failed to update meeting"));
      }
    },
  };
};

export default GoogleMeetVideoApiAdapter;
