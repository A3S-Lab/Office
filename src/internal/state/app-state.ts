export type OfficeNotificationTone = 'info' | 'success' | 'warning' | 'error';

export interface OfficeNotification {
  message: string;
  tone: OfficeNotificationTone;
}

export const OFFICE_NOTIFICATION_EVENT = 'a3s-office:notify';

export function showToast(
  message: string,
  tone: OfficeNotificationTone = 'info',
): void {
  if (typeof window === 'undefined') return;
  window.dispatchEvent(
    new CustomEvent<OfficeNotification>(OFFICE_NOTIFICATION_EVENT, {
      detail: { message, tone },
    }),
  );
}

export function subscribeOfficeNotifications(
  listener: (notification: OfficeNotification) => void,
): () => void {
  if (typeof window === 'undefined') return () => undefined;
  const receive = (event: Event) => {
    if (event instanceof CustomEvent)
      listener(event.detail as OfficeNotification);
  };
  window.addEventListener(OFFICE_NOTIFICATION_EVENT, receive);
  return () => window.removeEventListener(OFFICE_NOTIFICATION_EVENT, receive);
}
