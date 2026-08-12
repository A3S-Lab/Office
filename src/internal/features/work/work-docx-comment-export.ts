import type { ICommentOptions } from 'docx';
import type { AssignedDocxCommentThread } from './work-docx-note-comment-identity';

export function createDocxCommentRecords(
  threads: readonly AssignedDocxCommentThread[],
  docx: typeof import('docx'),
  commentIds: Map<string, number>,
): ICommentOptions[] {
  const records: ICommentOptions[] = [];
  for (const thread of threads) {
    const { comment, id } = thread;
    commentIds.set(comment.id, id);
    records.push({
      id,
      author: comment.author || 'A3S Work',
      initials: commentInitials(comment.author),
      date: commentDate(comment.date),
      resolved: Boolean(comment.resolved),
      children: [new docx.Paragraph({ text: comment.text || '（空批注）' })],
    });
    for (const assignedReply of thread.replies) {
      const { reply } = assignedReply;
      records.push({
        id: assignedReply.id,
        parentId: id,
        author: reply.author || 'A3S Work',
        initials: commentInitials(reply.author),
        date: commentDate(reply.date),
        children: [new docx.Paragraph({ text: reply.text || '（空回复）' })],
      });
    }
  }
  return records;
}

function commentDate(value: string): Date {
  const time = Date.parse(value);
  return Number.isFinite(time) ? new Date(time) : new Date();
}

function commentInitials(author: string): string {
  const parts = author.trim().split(/\s+/).filter(Boolean);
  if (!parts.length) return 'AW';
  return parts
    .slice(0, 2)
    .map((part) => Array.from(part)[0] ?? '')
    .join('')
    .toUpperCase();
}
