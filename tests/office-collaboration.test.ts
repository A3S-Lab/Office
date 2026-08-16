import { expect, test } from '@rstest/core';
import * as Y from 'yjs';
import {
  createOfficeCollaborationSession,
  createOfficeMarkdownCollaborationBinding,
  createOfficePdfCollaborationBinding,
  initializeOfficeMarkdownCollaboration,
  OfficeCollaborationError,
  type OfficeCollaborationMutationScope,
  type OfficeCollaborationOrigin,
  readOfficeCollaborationMetadata,
  readOfficeDocumentCollaboration,
  readOfficeMarkdownCollaboration,
  readOfficePdfCollaboration,
} from '../src/core';
import {
  BROWSER_PDF_FIXTURE_BASE64,
  NATIVE_PDF_CREATE_ANNOTATION_BASE64,
  NATIVE_PDF_DELETE_ANNOTATION_BASE64,
  NATIVE_PDF_UPDATE_ANNOTATION_BASE64,
} from './fixtures/native-pdf-annotation-updates';

const BROWSER_MARKDOWN_FIXTURE_BASE64 =
  'AQey8hkAKAETYTNzLm9mZmljZS5tZXRhZGF0YQhwcm90b2NvbAF3GGEzcy5vZmZpY2UuY29sbGFib3JhdGlvbigBE2Ezcy5vZmZpY2UubWV0YWRhdGEHdmVyc2lvbgF9ASgBE2Ezcy5vZmZpY2UubWV0YWRhdGEKYXJ0aWZhY3RJZAF3EGZpeHR1cmUtbWFya2Rvd24oARNhM3Mub2ZmaWNlLm1ldGFkYXRhBGtpbmQBdwhtYXJrZG93bigBE2Ezcy5vZmZpY2UubWV0YWRhdGELaW5pdGlhbGl6ZWQBeAgBIWEzcy5vZmZpY2UuYm9vdHN0cmFwLmluaXRpYWxpemVycwF3FjQyNDI0Mjpicm93c2VyLWZpeHR1cmUEARphM3Mub2ZmaWNlLm1hcmtkb3duLnNvdXJjZRUjIFNoYXJlZAoKWWpzIHRvIFlycy4A';
const NATIVE_MARKDOWN_REPLACE_BASE64 = 'AQGh9zYAhLLyGRoGQfCfmIBCAbLyGQEGFQ==';
const NATIVE_MARKDOWN_SPLICE_BASE64 =
  'AQGh9zYExKH3NgKh9zYDBPCfpoACsvIZAQYVofc2AQEC';
const BROWSER_DOCUMENT_FIXTURE_BASE64 =
  'AQ+y8hkAKAETYTNzLm9mZmljZS5tZXRhZGF0YQhwcm90b2NvbAF3GGEzcy5vZmZpY2UuY29sbGFib3JhdGlvbigBE2Ezcy5vZmZpY2UubWV0YWRhdGEHdmVyc2lvbgF9ASgBE2Ezcy5vZmZpY2UubWV0YWRhdGEKYXJ0aWZhY3RJZAF3EGZpeHR1cmUtZG9jdW1lbnQoARNhM3Mub2ZmaWNlLm1ldGFkYXRhBGtpbmQBdwhkb2N1bWVudCgBE2Ezcy5vZmZpY2UubWV0YWRhdGELaW5pdGlhbGl6ZWQBeAgBIWEzcy5vZmZpY2UuYm9vdHN0cmFwLmluaXRpYWxpemVycwF3FjQyNDI0Mjpicm93c2VyLWZpeHR1cmUHARthM3Mub2ZmaWNlLmRvY3VtZW50LmNvbnRlbnQDD2RvY3VtZW50U2VjdGlvbgcAsvIZBgMJcGFyYWdyYXBoBwCy8hkHBgQAsvIZCBBIZWxsbyDwn5iAIHdvcmxkKACy8hkHC3BhcmFncmFwaElkAXcIMDAwMDAwMDEoALLyGQcGdGV4dElkAXcIMDAwMDAwMDIoALLyGQYCaWQBdxJkb2N1bWVudC1zZWN0aW9uLTEoARthM3Mub2ZmaWNlLmRvY3VtZW50Lm9wdGlvbnMJcGFnZUNvbG9yAXcHI0Y4RkFGQygBG2Ezcy5vZmZpY2UuZG9jdW1lbnQub3B0aW9ucwx0cmFja0NoYW5nZXMBeAA=';
const NATIVE_DOCUMENT_REPLACE_BASE64 =
  'AQKi9zYAxLLyGRCy8hkRBPCfpoCosvIZGAF3CDAwMDAwMDAzAbLyGQIPAhgB';
const NATIVE_DOCUMENT_PAGE_COLOR_BASE64 =
  'AQGi9zYDqLLyGRoBdwcjMTAxODI4AbLyGQMPAhgBGgE=';
const NATIVE_DOCUMENT_TRACK_CHANGES_BASE64 = 'AAGy8hkDDwIYARoC';
const NATIVE_DOCUMENT_INSERT_TEMPORARY_BASE64 =
  'AQWi9zYEh7LyGQcDCXBhcmFncmFwaCgAovc2BAZ0ZXh0SWQBdwgwMDAwMDAxMSgAovc2BAtwYXJhZ3JhcGhJZAF3CDAwMDAwMDEwBwCi9zYEBgQAovc2BwlUZW1wb3JhcnkBsvIZAw8CGAEaAg==';
const NATIVE_DOCUMENT_INSERT_FINAL_BASE64 =
  'AQWi9zYRh6L3NgQDCXBhcmFncmFwaCgAovc2EQtwYXJhZ3JhcGhJZAF3CDAwMDAwMDEyKACi9zYRBnRleHRJZAF3CDAwMDAwMDEzBwCi9zYRBgQAovc2FBBOYXRpdmUgcGFyYWdyYXBoAbLyGQMPAhgBGgI=';
const NATIVE_DOCUMENT_DELETE_TEMPORARY_BASE64 = 'AAKy8hkDDwIYARoCovc2AQQN';
const NATIVE_PDF_SET_NAME_BASE64 = 'AQGj9zYAqLLyGQwBdwVHcmFjZQGy8hkCBAEMAQ==';
const NATIVE_PDF_SET_EMAIL_BASE64 =
  'AQSj9zYBKAEjYTNzLm9mZmljZS5wZGYuZm9ybS12YWx1ZXMucHJlc2VuY2UPQXBwbGljYW50LkVtYWlsAXgoASFhM3Mub2ZmaWNlLnBkZi5mb3JtLXZhbHVlcy5maWVsZHMoWyJBcHBsaWNhbnQuRW1haWwiLCJbXCJ2YWx1ZVwiLFwiaWRcIl0iXQF3D0FwcGxpY2FudC5FbWFpbIiy8hkNAXcPQXBwbGljYW50LkVtYWlsKAEhYTNzLm9mZmljZS5wZGYuZm9ybS12YWx1ZXMuZmllbGRzK1siQXBwbGljYW50LkVtYWlsIiwiW1widmFsdWVcIixcInZhbHVlXCJdIl0BdxJncmFjZUBleGFtcGxlLnRlc3QBsvIZAgQBDAE=';
const NATIVE_PDF_PROPOSE_REDACTION_BASE64 =
  'AQqj9zYAKAErYTNzLm9mZmljZS5wZGYucmVkYWN0aW9uLXByb3Bvc2Fscy5wcmVzZW5jZRJyZWRhY3Rpb24tbmF0aXZlLTEBeCgBKWEzcy5vZmZpY2UucGRmLnJlZGFjdGlvbi1wcm9wb3NhbHMuZmllbGRzK1sicmVkYWN0aW9uLW5hdGl2ZS0xIiwiW1widmFsdWVcIixcImlkXCJdIl0BdxJyZWRhY3Rpb24tbmF0aXZlLTEoASlhM3Mub2ZmaWNlLnBkZi5yZWRhY3Rpb24tcHJvcG9zYWxzLmZpZWxkczJbInJlZGFjdGlvbi1uYXRpdmUtMSIsIltcInZhbHVlXCIsXCJwYWdlSW5kZXhcIl0iXQF9ASgBKWEzcy5vZmZpY2UucGRmLnJlZGFjdGlvbi1wcm9wb3NhbHMuZmllbGRzM1sicmVkYWN0aW9uLW5hdGl2ZS0xIiwiW1widmFsdWVcIixcInByb3Bvc2VkQXRcIl0iXQF3GDIwMjYtMDgtMTVUMDM6MDA6MDAuMDAwWigBKWEzcy5vZmZpY2UucGRmLnJlZGFjdGlvbi1wcm9wb3NhbHMuZmllbGRzM1sicmVkYWN0aW9uLW5hdGl2ZS0xIiwiW1widmFsdWVcIixcInByb3Bvc2VkQnlcIl0iXQF3C2FnZW50LWFscGhhKAEpYTNzLm9mZmljZS5wZGYucmVkYWN0aW9uLXByb3Bvc2Fscy5maWVsZHMvWyJyZWRhY3Rpb24tbmF0aXZlLTEiLCJbXCJ2YWx1ZVwiLFwicmVhc29uXCJdIl0Bdw1QZXJzb25hbCBkYXRhKAEpYTNzLm9mZmljZS5wZGYucmVkYWN0aW9uLXByb3Bvc2Fscy5maWVsZHMuWyJyZWRhY3Rpb24tbmF0aXZlLTEiLCJbXCJ2YWx1ZVwiLFwicmVjdHNcIl0iXQF1AnYEBXJpZ2h0fEKhgAADdG9wfEGiAAAGYm90dG9tfEIiAAAEbGVmdHxBKAAAdgQFcmlnaHR9uAEEbGVmdH2aAQZib3R0b218QiIAAAN0b3B8QaIAACgBKWEzcy5vZmZpY2UucGRmLnJlZGFjdGlvbi1wcm9wb3NhbHMuZmllbGRzLVsicmVkYWN0aW9uLW5hdGl2ZS0xIiwiW1widmFsdWVcIixcInRleHRcIl0iXQF3DEFjY291bnQgMTIzNAgBKGEzcy5vZmZpY2UucGRmLnJlZGFjdGlvbi1wcm9wb3NhbHMub3JkZXIBdxJyZWRhY3Rpb24tbmF0aXZlLTEIARxhM3Mub2ZmaWNlLnBkZi5yZWNvcmQtY2xhaW1zAXf0AnsiZmluZ2VycHJpbnQiOiJ7XCJpZFwiOlwicmVkYWN0aW9uLW5hdGl2ZS0xXCIsXCJwYWdlSW5kZXhcIjoxLFwicHJvcG9zZWRBdFwiOlwiMjAyNi0wOC0xNVQwMzowMDowMC4wMDBaXCIsXCJwcm9wb3NlZEJ5XCI6XCJhZ2VudC1hbHBoYVwiLFwicmVhc29uXCI6XCJQZXJzb25hbCBkYXRhXCIsXCJyZWN0c1wiOlt7XCJib3R0b21cIjo0MC41LFwibGVmdFwiOjEwLjUsXCJyaWdodFwiOjgwLjc1LFwidG9wXCI6MjAuMjV9LHtcImJvdHRvbVwiOjQwLjUsXCJsZWZ0XCI6OTAsXCJyaWdodFwiOjEyMCxcInRvcFwiOjIwLjI1fV0sXCJ0ZXh0XCI6XCJBY2NvdW50IDEyMzRcIn0iLCJpZCI6InJlZGFjdGlvbi1uYXRpdmUtMSIsImtpbmQiOiJyZWRhY3Rpb24ifQGy8hkBBAE=';
const NATIVE_PDF_DECIDE_REVIEW_BASE64 =
  'AQmj9zYKKAEoYTNzLm9mZmljZS5wZGYucmV2aWV3LWRlY2lzaW9ucy5wcmVzZW5jZRFkZWNpc2lvbi1uYXRpdmUtMQF4KAEmYTNzLm9mZmljZS5wZGYucmV2aWV3LWRlY2lzaW9ucy5maWVsZHMvWyJkZWNpc2lvbi1uYXRpdmUtMSIsIltcInZhbHVlXCIsXCJhY3RvcklkXCJdIl0BdwthZ2VudC1hbHBoYSgBJmEzcy5vZmZpY2UucGRmLnJldmlldy1kZWNpc2lvbnMuZmllbGRzMVsiZGVjaXNpb24tbmF0aXZlLTEiLCJbXCJ2YWx1ZVwiLFwiY3JlYXRlZEF0XCJdIl0BdxgyMDI2LTA4LTE1VDAzOjA1OjAwLjAwMFooASZhM3Mub2ZmaWNlLnBkZi5yZXZpZXctZGVjaXNpb25zLmZpZWxkczBbImRlY2lzaW9uLW5hdGl2ZS0xIiwiW1widmFsdWVcIixcImRlY2lzaW9uXCJdIl0BdwdhcHByb3ZlKAEmYTNzLm9mZmljZS5wZGYucmV2aWV3LWRlY2lzaW9ucy5maWVsZHMqWyJkZWNpc2lvbi1uYXRpdmUtMSIsIltcInZhbHVlXCIsXCJpZFwiXSJdAXcRZGVjaXNpb24tbmF0aXZlLTEoASZhM3Mub2ZmaWNlLnBkZi5yZXZpZXctZGVjaXNpb25zLmZpZWxkczBbImRlY2lzaW9uLW5hdGl2ZS0xIiwiW1widmFsdWVcIixcInRhcmdldElkXCJdIl0BdxJyZWRhY3Rpb24tbmF0aXZlLTEoASZhM3Mub2ZmaWNlLnBkZi5yZXZpZXctZGVjaXNpb25zLmZpZWxkczJbImRlY2lzaW9uLW5hdGl2ZS0xIiwiW1widmFsdWVcIixcInRhcmdldEtpbmRcIl0iXQF3CXJlZGFjdGlvbggBJWEzcy5vZmZpY2UucGRmLnJldmlldy1kZWNpc2lvbnMub3JkZXIBdxFkZWNpc2lvbi1uYXRpdmUtMYij9zYJAXeDAnsiZmluZ2VycHJpbnQiOiJ7XCJhY3RvcklkXCI6XCJhZ2VudC1hbHBoYVwiLFwiY3JlYXRlZEF0XCI6XCIyMDI2LTA4LTE1VDAzOjA1OjAwLjAwMFpcIixcImRlY2lzaW9uXCI6XCJhcHByb3ZlXCIsXCJpZFwiOlwiZGVjaXNpb24tbmF0aXZlLTFcIixcInRhcmdldElkXCI6XCJyZWRhY3Rpb24tbmF0aXZlLTFcIixcInRhcmdldEtpbmRcIjpcInJlZGFjdGlvblwifSIsImlkIjoiZGVjaXNpb24tbmF0aXZlLTEiLCJraW5kIjoicmV2aWV3LWRlY2lzaW9uIn0BsvIZAQQB';
const NATIVE_PDF_PROPOSE_PAGE_ROTATION_BASE64 =
  'AQmj9zYAKAEnYTNzLm9mZmljZS5wZGYucGFnZS1vcGVyYXRpb25zLnByZXNlbmNlF3BhZ2Utb3BlcmF0aW9uLXJvdGF0ZS0xAXgoASVhM3Mub2ZmaWNlLnBkZi5wYWdlLW9wZXJhdGlvbnMuZmllbGRzNVsicGFnZS1vcGVyYXRpb24tcm90YXRlLTEiLCJbXCJ2YWx1ZVwiLFwiZGVncmVlc1wiXSJdAX2aASgBJWEzcy5vZmZpY2UucGRmLnBhZ2Utb3BlcmF0aW9ucy5maWVsZHMwWyJwYWdlLW9wZXJhdGlvbi1yb3RhdGUtMSIsIltcInZhbHVlXCIsXCJpZFwiXSJdAXcXcGFnZS1vcGVyYXRpb24tcm90YXRlLTEoASVhM3Mub2ZmaWNlLnBkZi5wYWdlLW9wZXJhdGlvbnMuZmllbGRzMlsicGFnZS1vcGVyYXRpb24tcm90YXRlLTEiLCJbXCJ2YWx1ZVwiLFwia2luZFwiXSJdAXcGcm90YXRlKAElYTNzLm9mZmljZS5wZGYucGFnZS1vcGVyYXRpb25zLmZpZWxkczlbInBhZ2Utb3BlcmF0aW9uLXJvdGF0ZS0xIiwiW1widmFsdWVcIixcInBhZ2VJbmRpY2VzXCJdIl0BdQJ9AH0CKAElYTNzLm9mZmljZS5wZGYucGFnZS1vcGVyYXRpb25zLmZpZWxkczhbInBhZ2Utb3BlcmF0aW9uLXJvdGF0ZS0xIiwiW1widmFsdWVcIixcInByb3Bvc2VkQXRcIl0iXQF3GDIwMjYtMDgtMTVUMDc6MDA6MDAuMDAwWigBJWEzcy5vZmZpY2UucGRmLnBhZ2Utb3BlcmF0aW9ucy5maWVsZHM4WyJwYWdlLW9wZXJhdGlvbi1yb3RhdGUtMSIsIltcInZhbHVlXCIsXCJwcm9wb3NlZEJ5XCJdIl0BdwthZ2VudC1hbHBoYQgBJGEzcy5vZmZpY2UucGRmLnBhZ2Utb3BlcmF0aW9ucy5vcmRlcgF3F3BhZ2Utb3BlcmF0aW9uLXJvdGF0ZS0xCAEcYTNzLm9mZmljZS5wZGYucmVjb3JkLWNsYWltcwF38QF7ImZpbmdlcnByaW50Ijoie1wiZGVncmVlc1wiOjkwLFwiaWRcIjpcInBhZ2Utb3BlcmF0aW9uLXJvdGF0ZS0xXCIsXCJraW5kXCI6XCJyb3RhdGVcIixcInBhZ2VJbmRpY2VzXCI6WzAsMl0sXCJwcm9wb3NlZEF0XCI6XCIyMDI2LTA4LTE1VDA3OjAwOjAwLjAwMFpcIixcInByb3Bvc2VkQnlcIjpcImFnZW50LWFscGhhXCJ9IiwiaWQiOiJwYWdlLW9wZXJhdGlvbi1yb3RhdGUtMSIsImtpbmQiOiJwYWdlLW9wZXJhdGlvbiJ9AbLyGQEEAQ==';
const NATIVE_PDF_PROPOSE_PAGE_DELETION_BASE64 =
  'AQij9zYJKAEnYTNzLm9mZmljZS5wZGYucGFnZS1vcGVyYXRpb25zLnByZXNlbmNlF3BhZ2Utb3BlcmF0aW9uLWRlbGV0ZS0xAXgoASVhM3Mub2ZmaWNlLnBkZi5wYWdlLW9wZXJhdGlvbnMuZmllbGRzMFsicGFnZS1vcGVyYXRpb24tZGVsZXRlLTEiLCJbXCJ2YWx1ZVwiLFwiaWRcIl0iXQF3F3BhZ2Utb3BlcmF0aW9uLWRlbGV0ZS0xKAElYTNzLm9mZmljZS5wZGYucGFnZS1vcGVyYXRpb25zLmZpZWxkczJbInBhZ2Utb3BlcmF0aW9uLWRlbGV0ZS0xIiwiW1widmFsdWVcIixcImtpbmRcIl0iXQF3BmRlbGV0ZSgBJWEzcy5vZmZpY2UucGRmLnBhZ2Utb3BlcmF0aW9ucy5maWVsZHM5WyJwYWdlLW9wZXJhdGlvbi1kZWxldGUtMSIsIltcInZhbHVlXCIsXCJwYWdlSW5kaWNlc1wiXSJdAXUBfQIoASVhM3Mub2ZmaWNlLnBkZi5wYWdlLW9wZXJhdGlvbnMuZmllbGRzOFsicGFnZS1vcGVyYXRpb24tZGVsZXRlLTEiLCJbXCJ2YWx1ZVwiLFwicHJvcG9zZWRBdFwiXSJdAXcYMjAyNi0wOC0xNVQwNzowMTowMC4wMDBaKAElYTNzLm9mZmljZS5wZGYucGFnZS1vcGVyYXRpb25zLmZpZWxkczhbInBhZ2Utb3BlcmF0aW9uLWRlbGV0ZS0xIiwiW1widmFsdWVcIixcInByb3Bvc2VkQnlcIl0iXQF3C2FnZW50LWFscGhhiKP3NgcBdxdwYWdlLW9wZXJhdGlvbi1kZWxldGUtMYij9zYIAXfgAXsiZmluZ2VycHJpbnQiOiJ7XCJpZFwiOlwicGFnZS1vcGVyYXRpb24tZGVsZXRlLTFcIixcImtpbmRcIjpcImRlbGV0ZVwiLFwicGFnZUluZGljZXNcIjpbMl0sXCJwcm9wb3NlZEF0XCI6XCIyMDI2LTA4LTE1VDA3OjAxOjAwLjAwMFpcIixcInByb3Bvc2VkQnlcIjpcImFnZW50LWFscGhhXCJ9IiwiaWQiOiJwYWdlLW9wZXJhdGlvbi1kZWxldGUtMSIsImtpbmQiOiJwYWdlLW9wZXJhdGlvbiJ9AbLyGQEEAQ==';
const NATIVE_PDF_PROPOSE_PAGE_REORDER_BASE64 =
  'AQij9zYRKAEnYTNzLm9mZmljZS5wZGYucGFnZS1vcGVyYXRpb25zLnByZXNlbmNlGHBhZ2Utb3BlcmF0aW9uLXJlb3JkZXItMQF4KAElYTNzLm9mZmljZS5wZGYucGFnZS1vcGVyYXRpb25zLmZpZWxkczFbInBhZ2Utb3BlcmF0aW9uLXJlb3JkZXItMSIsIltcInZhbHVlXCIsXCJpZFwiXSJdAXcYcGFnZS1vcGVyYXRpb24tcmVvcmRlci0xKAElYTNzLm9mZmljZS5wZGYucGFnZS1vcGVyYXRpb25zLmZpZWxkczNbInBhZ2Utb3BlcmF0aW9uLXJlb3JkZXItMSIsIltcInZhbHVlXCIsXCJraW5kXCJdIl0BdwdyZW9yZGVyKAElYTNzLm9mZmljZS5wZGYucGFnZS1vcGVyYXRpb25zLmZpZWxkczhbInBhZ2Utb3BlcmF0aW9uLXJlb3JkZXItMSIsIltcInZhbHVlXCIsXCJwYWdlT3JkZXJcIl0iXQF1A30CfQB9ASgBJWEzcy5vZmZpY2UucGRmLnBhZ2Utb3BlcmF0aW9ucy5maWVsZHM5WyJwYWdlLW9wZXJhdGlvbi1yZW9yZGVyLTEiLCJbXCJ2YWx1ZVwiLFwicHJvcG9zZWRBdFwiXSJdAXcYMjAyNi0wOC0xNVQwNzowMjowMC4wMDBaKAElYTNzLm9mZmljZS5wZGYucGFnZS1vcGVyYXRpb25zLmZpZWxkczlbInBhZ2Utb3BlcmF0aW9uLXJlb3JkZXItMSIsIltcInZhbHVlXCIsXCJwcm9wb3NlZEJ5XCJdIl0BdwthZ2VudC1hbHBoYYij9zYPAXcYcGFnZS1vcGVyYXRpb24tcmVvcmRlci0xiKP3NhABd+UBeyJmaW5nZXJwcmludCI6IntcImlkXCI6XCJwYWdlLW9wZXJhdGlvbi1yZW9yZGVyLTFcIixcImtpbmRcIjpcInJlb3JkZXJcIixcInBhZ2VPcmRlclwiOlsyLDAsMV0sXCJwcm9wb3NlZEF0XCI6XCIyMDI2LTA4LTE1VDA3OjAyOjAwLjAwMFpcIixcInByb3Bvc2VkQnlcIjpcImFnZW50LWFscGhhXCJ9IiwiaWQiOiJwYWdlLW9wZXJhdGlvbi1yZW9yZGVyLTEiLCJraW5kIjoicGFnZS1vcGVyYXRpb24ifQGy8hkBBAE=';
const NATIVE_PDF_DECIDE_PAGE_OPERATION_BASE64 =
  'AQmj9zYZKAEoYTNzLm9mZmljZS5wZGYucmV2aWV3LWRlY2lzaW9ucy5wcmVzZW5jZRlkZWNpc2lvbi1wYWdlLW9wZXJhdGlvbi0xAXgoASZhM3Mub2ZmaWNlLnBkZi5yZXZpZXctZGVjaXNpb25zLmZpZWxkczdbImRlY2lzaW9uLXBhZ2Utb3BlcmF0aW9uLTEiLCJbXCJ2YWx1ZVwiLFwiYWN0b3JJZFwiXSJdAXcLYWdlbnQtYWxwaGEoASZhM3Mub2ZmaWNlLnBkZi5yZXZpZXctZGVjaXNpb25zLmZpZWxkczlbImRlY2lzaW9uLXBhZ2Utb3BlcmF0aW9uLTEiLCJbXCJ2YWx1ZVwiLFwiY3JlYXRlZEF0XCJdIl0BdxgyMDI2LTA4LTE1VDA3OjA1OjAwLjAwMFooASZhM3Mub2ZmaWNlLnBkZi5yZXZpZXctZGVjaXNpb25zLmZpZWxkczhbImRlY2lzaW9uLXBhZ2Utb3BlcmF0aW9uLTEiLCJbXCJ2YWx1ZVwiLFwiZGVjaXNpb25cIl0iXQF3B2FwcHJvdmUoASZhM3Mub2ZmaWNlLnBkZi5yZXZpZXctZGVjaXNpb25zLmZpZWxkczJbImRlY2lzaW9uLXBhZ2Utb3BlcmF0aW9uLTEiLCJbXCJ2YWx1ZVwiLFwiaWRcIl0iXQF3GWRlY2lzaW9uLXBhZ2Utb3BlcmF0aW9uLTEoASZhM3Mub2ZmaWNlLnBkZi5yZXZpZXctZGVjaXNpb25zLmZpZWxkczhbImRlY2lzaW9uLXBhZ2Utb3BlcmF0aW9uLTEiLCJbXCJ2YWx1ZVwiLFwidGFyZ2V0SWRcIl0iXQF3F3BhZ2Utb3BlcmF0aW9uLWRlbGV0ZS0xKAEmYTNzLm9mZmljZS5wZGYucmV2aWV3LWRlY2lzaW9ucy5maWVsZHM6WyJkZWNpc2lvbi1wYWdlLW9wZXJhdGlvbi0xIiwiW1widmFsdWVcIixcInRhcmdldEtpbmRcIl0iXQF3DnBhZ2Utb3BlcmF0aW9uCAElYTNzLm9mZmljZS5wZGYucmV2aWV3LWRlY2lzaW9ucy5vcmRlcgF3GWRlY2lzaW9uLXBhZ2Utb3BlcmF0aW9uLTGIo/c2GAF3nQJ7ImZpbmdlcnByaW50Ijoie1wiYWN0b3JJZFwiOlwiYWdlbnQtYWxwaGFcIixcImNyZWF0ZWRBdFwiOlwiMjAyNi0wOC0xNVQwNzowNTowMC4wMDBaXCIsXCJkZWNpc2lvblwiOlwiYXBwcm92ZVwiLFwiaWRcIjpcImRlY2lzaW9uLXBhZ2Utb3BlcmF0aW9uLTFcIixcInRhcmdldElkXCI6XCJwYWdlLW9wZXJhdGlvbi1kZWxldGUtMVwiLFwidGFyZ2V0S2luZFwiOlwicGFnZS1vcGVyYXRpb25cIn0iLCJpZCI6ImRlY2lzaW9uLXBhZ2Utb3BlcmF0aW9uLTEiLCJraW5kIjoicmV2aWV3LWRlY2lzaW9uIn0BsvIZAQQB';

test('applies UTF-16-safe native typed Markdown updates in browser Yjs', () => {
  const document = new Y.Doc();
  for (const encoded of [
    BROWSER_MARKDOWN_FIXTURE_BASE64,
    NATIVE_MARKDOWN_REPLACE_BASE64,
    NATIVE_MARKDOWN_SPLICE_BASE64,
  ]) {
    Y.applyUpdate(document, decodeBase64(encoded));
  }

  expect(document.getText('a3s.office.markdown.source').toString()).toBe(
    'A🦀B',
  );
});

test('applies native typed Document text, paragraph, and option updates in browser Yjs', () => {
  const browserDocument = new Y.Doc();
  const nativeDocument = new Y.Doc();
  for (const document of [browserDocument, nativeDocument]) {
    Y.applyUpdate(document, decodeBase64(BROWSER_DOCUMENT_FIXTURE_BASE64));
  }
  const section = browserDocument
    .getXmlFragment('a3s.office.document.content')
    .get(0);
  const paragraph = section instanceof Y.XmlElement ? section.get(0) : null;
  const text = paragraph instanceof Y.XmlElement ? paragraph.get(0) : null;
  if (!(text instanceof Y.XmlText)) {
    throw new Error('Expected the browser Document fixture text node.');
  }
  browserDocument.transact(() => {
    text.insert(text.length, ' from browser');
    const browserParagraph = new Y.XmlElement('paragraph');
    browserParagraph.setAttribute('paragraphId', '00000020');
    browserParagraph.setAttribute('textId', '00000021');
    const browserText = new Y.XmlText();
    browserText.insert(0, 'Browser paragraph');
    browserParagraph.insert(0, [browserText]);
    if (!(section instanceof Y.XmlElement)) {
      throw new Error('Expected the browser Document fixture section.');
    }
    section.insert(1, [browserParagraph]);
  });
  for (const encoded of [
    NATIVE_DOCUMENT_REPLACE_BASE64,
    NATIVE_DOCUMENT_PAGE_COLOR_BASE64,
    NATIVE_DOCUMENT_TRACK_CHANGES_BASE64,
    NATIVE_DOCUMENT_INSERT_TEMPORARY_BASE64,
    NATIVE_DOCUMENT_INSERT_FINAL_BASE64,
    NATIVE_DOCUMENT_DELETE_TEMPORARY_BASE64,
  ]) {
    Y.applyUpdate(nativeDocument, decodeBase64(encoded));
  }
  exchangeUpdates(browserDocument, nativeDocument);

  const contents = [browserDocument, nativeDocument].map((document) =>
    readOfficeDocumentCollaboration(
      createOfficeCollaborationSession({
        artifactId: 'fixture-document',
        document,
        kind: 'document',
      }),
    ),
  );
  for (const content of contents) {
    expect(content.html).toContain('Hello 🦀 world from browser');
    expect(content.html).toContain('Native paragraph');
    expect(content.html).toContain('Browser paragraph');
    expect(content.html).not.toContain('Temporary');
    expect(content.pageColor).toBe('#101828');
    expect(content.trackChanges).toBeUndefined();
    const paragraphs = content.model?.root.content?.[0]?.content ?? [];
    expect(
      paragraphs.map((node) => node.content?.map((item) => item.text).join('')),
    ).toEqual(
      expect.arrayContaining([
        'Hello 🦀 world from browser',
        'Native paragraph',
        'Browser paragraph',
      ]),
    );
    expect(paragraphs).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          attrs: expect.objectContaining({
            paragraphId: '00000001',
            textId: '00000003',
          }),
        }),
        expect.objectContaining({
          attrs: expect.objectContaining({
            paragraphId: '00000012',
            textId: '00000013',
          }),
        }),
        expect.objectContaining({
          attrs: expect.objectContaining({
            paragraphId: '00000020',
            textId: '00000021',
          }),
        }),
      ]),
    );
  }
  expect(contents[0]?.html).toBe(contents[1]?.html);
});

test('applies native typed PDF form updates and converges with a browser edit', () => {
  const browserDocument = new Y.Doc();
  const nativeDocument = new Y.Doc();
  const reorderedDocument = new Y.Doc();
  for (const document of [browserDocument, nativeDocument, reorderedDocument]) {
    Y.applyUpdate(document, decodeBase64(BROWSER_PDF_FIXTURE_BASE64));
  }
  const browserSession = createOfficeCollaborationSession({
    artifactId: 'fixture-pdf',
    document: browserDocument,
    kind: 'pdf',
  });
  const browserBinding = createOfficePdfCollaborationBinding(browserSession);
  const before = browserBinding.content();
  browserBinding.replace(before, {
    ...before,
    formValues: [
      ...before.formValues,
      { id: 'Applicant.Phone', value: '+1 555 0100' },
    ],
  });

  for (const encoded of [
    NATIVE_PDF_SET_NAME_BASE64,
    NATIVE_PDF_SET_EMAIL_BASE64,
  ]) {
    Y.applyUpdate(nativeDocument, decodeBase64(encoded));
  }
  for (const encoded of [
    NATIVE_PDF_SET_EMAIL_BASE64,
    NATIVE_PDF_SET_NAME_BASE64,
    NATIVE_PDF_SET_EMAIL_BASE64,
  ]) {
    Y.applyUpdate(reorderedDocument, decodeBase64(encoded));
  }
  exchangeUpdates(browserDocument, nativeDocument);
  exchangeUpdates(browserDocument, reorderedDocument);

  const contents = [browserDocument, nativeDocument, reorderedDocument].map(
    (document) =>
      readOfficePdfCollaboration(
        createOfficeCollaborationSession({
          artifactId: 'fixture-pdf',
          document,
          kind: 'pdf',
        }),
      ),
  );
  expect(contents[0]).toEqual(contents[1]);
  expect(contents[0]).toEqual(contents[2]);
  expect(contents[0]?.formValues).toEqual(
    expect.arrayContaining([
      { id: 'Applicant.Name', value: 'Grace' },
      { id: 'Applicant.Email', value: 'grace@example.test' },
      { id: 'Applicant.Phone', value: '+1 555 0100' },
    ]),
  );
});

test('merges native PDF annotation mutations with browser leaves across reordered delivery', () => {
  const orderedDocument = new Y.Doc();
  const reorderedDocument = new Y.Doc();
  for (const document of [orderedDocument, reorderedDocument]) {
    Y.applyUpdate(document, decodeBase64(BROWSER_PDF_FIXTURE_BASE64));
    Y.applyUpdate(document, decodeBase64(NATIVE_PDF_CREATE_ANNOTATION_BASE64));
    const session = createOfficeCollaborationSession({
      artifactId: 'fixture-pdf',
      document,
      kind: 'pdf',
    });
    const binding = createOfficePdfCollaborationBinding(session);
    const before = binding.content();
    binding.replace(before, {
      ...before,
      annotations: before.annotations.map((record) =>
        record.id === 'annotation-native-interop'
          ? {
              ...record,
              annotation: {
                ...record.annotation,
                contents: 'Browser note',
              },
            }
          : record,
      ),
    });
  }

  for (const encoded of [
    NATIVE_PDF_UPDATE_ANNOTATION_BASE64,
    NATIVE_PDF_DELETE_ANNOTATION_BASE64,
  ]) {
    Y.applyUpdate(orderedDocument, decodeBase64(encoded));
  }
  for (const encoded of [
    NATIVE_PDF_DELETE_ANNOTATION_BASE64,
    NATIVE_PDF_UPDATE_ANNOTATION_BASE64,
    NATIVE_PDF_UPDATE_ANNOTATION_BASE64,
  ]) {
    Y.applyUpdate(reorderedDocument, decodeBase64(encoded));
  }

  const contents = [orderedDocument, reorderedDocument].map((document) =>
    readOfficePdfCollaboration(
      createOfficeCollaborationSession({
        artifactId: 'fixture-pdf',
        document,
        kind: 'pdf',
      }),
    ),
  );
  for (const content of contents) {
    const annotation = content.annotations.find(
      ({ id }) => id === 'annotation-native-interop',
    );
    expect(annotation).toMatchObject({
      id: 'annotation-native-interop',
      pageIndex: 1,
      source: 'created',
      deleted: true,
      annotation: {
        id: 'annotation-native-interop',
        pageIndex: 1,
        type: 9,
        color: '#ff0000',
        contents: 'Browser note',
      },
    });
  }
});

test('applies native PDF review records, claims, and reordered duplicate delivery in browser Yjs', () => {
  const browserDocument = new Y.Doc();
  const nativeDocument = new Y.Doc();
  const reorderedDocument = new Y.Doc();
  for (const document of [browserDocument, nativeDocument, reorderedDocument]) {
    Y.applyUpdate(document, decodeBase64(BROWSER_PDF_FIXTURE_BASE64));
  }

  const browserSession = createOfficeCollaborationSession({
    actor: { id: 'browser-reviewer', kind: 'human', name: 'Browser reviewer' },
    artifactId: 'fixture-pdf',
    document: browserDocument,
    kind: 'pdf',
  });
  const browserBinding = createOfficePdfCollaborationBinding(browserSession);
  const before = browserBinding.content();
  browserBinding.replace(before, {
    ...before,
    redactionProposals: [
      ...before.redactionProposals,
      {
        id: 'redaction-browser-1',
        pageIndex: 0,
        rects: [{ left: 5, top: 6, right: 35, bottom: 16 }],
        proposedBy: 'browser-reviewer',
        proposedAt: '2026-08-15T02:55:00.000Z',
        reason: 'Browser proposal',
      },
    ],
  });

  for (const encoded of [
    NATIVE_PDF_PROPOSE_REDACTION_BASE64,
    NATIVE_PDF_DECIDE_REVIEW_BASE64,
  ]) {
    Y.applyUpdate(nativeDocument, decodeBase64(encoded));
  }
  for (const encoded of [
    NATIVE_PDF_DECIDE_REVIEW_BASE64,
    NATIVE_PDF_PROPOSE_REDACTION_BASE64,
    NATIVE_PDF_DECIDE_REVIEW_BASE64,
  ]) {
    Y.applyUpdate(reorderedDocument, decodeBase64(encoded));
  }

  exchangeUpdates(browserDocument, nativeDocument);
  exchangeUpdates(browserDocument, reorderedDocument);
  const contents = [browserDocument, nativeDocument, reorderedDocument].map(
    (document) =>
      readOfficePdfCollaboration(
        createOfficeCollaborationSession({
          artifactId: 'fixture-pdf',
          document,
          kind: 'pdf',
        }),
      ),
  );
  expect(contents[0]).toEqual(contents[1]);
  expect(contents[0]).toEqual(contents[2]);
  expect(contents[0]?.redactionProposals).toEqual(
    expect.arrayContaining([
      expect.objectContaining({
        id: 'redaction-native-1',
        pageIndex: 1,
        proposedBy: 'agent-alpha',
        proposedAt: '2026-08-15T03:00:00.000Z',
        reason: 'Personal data',
        text: 'Account 1234',
      }),
      expect.objectContaining({
        id: 'redaction-browser-1',
        proposedBy: 'browser-reviewer',
      }),
    ]),
  );
  expect(contents[0]?.reviewDecisions).toContainEqual({
    id: 'decision-native-1',
    targetKind: 'redaction',
    targetId: 'redaction-native-1',
    decision: 'approve',
    actorId: 'agent-alpha',
    createdAt: '2026-08-15T03:05:00.000Z',
  });
});

test('applies native PDF page-operation proposals and decisions in browser Yjs', () => {
  const browserDocument = new Y.Doc();
  const nativeDocument = new Y.Doc();
  const reorderedDocument = new Y.Doc();
  for (const document of [browserDocument, nativeDocument, reorderedDocument]) {
    Y.applyUpdate(document, decodeBase64(BROWSER_PDF_FIXTURE_BASE64));
  }

  const browserBinding = createOfficePdfCollaborationBinding(
    createOfficeCollaborationSession({
      actor: {
        id: 'browser-page-reviewer',
        kind: 'human',
        name: 'Browser page reviewer',
      },
      artifactId: 'fixture-pdf',
      document: browserDocument,
      kind: 'pdf',
    }),
  );
  const before = browserBinding.content();
  browserBinding.replace(before, {
    ...before,
    pageOperations: [
      ...before.pageOperations,
      {
        id: 'page-operation-browser-1',
        kind: 'rotate',
        pageIndices: [1],
        degrees: 180,
        proposedBy: 'browser-page-reviewer',
        proposedAt: '2026-08-15T06:55:00.000Z',
      },
    ],
  });

  const nativeUpdates = [
    NATIVE_PDF_PROPOSE_PAGE_ROTATION_BASE64,
    NATIVE_PDF_PROPOSE_PAGE_DELETION_BASE64,
    NATIVE_PDF_PROPOSE_PAGE_REORDER_BASE64,
    NATIVE_PDF_DECIDE_PAGE_OPERATION_BASE64,
  ];
  for (const encoded of nativeUpdates) {
    Y.applyUpdate(nativeDocument, decodeBase64(encoded));
  }
  for (const encoded of [
    NATIVE_PDF_DECIDE_PAGE_OPERATION_BASE64,
    NATIVE_PDF_PROPOSE_PAGE_REORDER_BASE64,
    NATIVE_PDF_PROPOSE_PAGE_DELETION_BASE64,
    NATIVE_PDF_PROPOSE_PAGE_ROTATION_BASE64,
    NATIVE_PDF_PROPOSE_PAGE_DELETION_BASE64,
  ]) {
    Y.applyUpdate(reorderedDocument, decodeBase64(encoded));
  }

  exchangeUpdates(browserDocument, nativeDocument);
  exchangeUpdates(browserDocument, reorderedDocument);
  const contents = [browserDocument, nativeDocument, reorderedDocument].map(
    (document) =>
      readOfficePdfCollaboration(
        createOfficeCollaborationSession({
          artifactId: 'fixture-pdf',
          document,
          kind: 'pdf',
        }),
      ),
  );
  expect(contents[0]).toEqual(contents[1]);
  expect(contents[0]).toEqual(contents[2]);
  expect(contents[0]?.pageOperations).toEqual(
    expect.arrayContaining([
      {
        id: 'page-operation-rotate-1',
        kind: 'rotate',
        pageIndices: [0, 2],
        degrees: 90,
        proposedBy: 'agent-alpha',
        proposedAt: '2026-08-15T07:00:00.000Z',
      },
      {
        id: 'page-operation-delete-1',
        kind: 'delete',
        pageIndices: [2],
        proposedBy: 'agent-alpha',
        proposedAt: '2026-08-15T07:01:00.000Z',
      },
      {
        id: 'page-operation-reorder-1',
        kind: 'reorder',
        pageOrder: [2, 0, 1],
        proposedBy: 'agent-alpha',
        proposedAt: '2026-08-15T07:02:00.000Z',
      },
      expect.objectContaining({
        id: 'page-operation-browser-1',
        proposedBy: 'browser-page-reviewer',
      }),
    ]),
  );
  expect(contents[0]?.reviewDecisions).toContainEqual({
    id: 'decision-page-operation-1',
    targetKind: 'page-operation',
    targetId: 'page-operation-delete-1',
    decision: 'approve',
    actorId: 'agent-alpha',
    createdAt: '2026-08-15T07:05:00.000Z',
  });
});

test('rejects forged transaction and binding origins at runtime', () => {
  const session = createOfficeCollaborationSession({
    artifactId: 'notes-invalid-origin',
    kind: 'markdown',
  });
  initializeOfficeMarkdownCollaboration(session, {
    type: 'markdown',
    markdown: 'Shared',
  });
  const invalidOrigin = {
    protocol: 'a3s.office.collaboration',
    kind: 'untrusted',
  } as unknown as OfficeCollaborationOrigin;

  expect(() => session.transact(() => undefined, invalidOrigin)).toThrow(
    /origin kind 'untrusted' is invalid/,
  );
  expect(() =>
    session.transact(
      () => undefined,
      undefined,
      'untrusted' as OfficeCollaborationMutationScope,
    ),
  ).toThrow(/mutation scope 'untrusted' is invalid/);
  expect(() =>
    createOfficeMarkdownCollaborationBinding(session, {
      origin: invalidOrigin,
    }),
  ).toThrow(/origin kind 'untrusted' is invalid/);
});

function decodeBase64(value: string): Uint8Array {
  return Uint8Array.from(atob(value), (character) => character.charCodeAt(0));
}

test('initializes one versioned Markdown collaboration document', () => {
  const document = new Y.Doc();
  const session = createOfficeCollaborationSession({
    artifactId: 'notes-1',
    document,
    kind: 'markdown',
  });

  expect(
    initializeOfficeMarkdownCollaboration(session, {
      type: 'markdown',
      markdown: '# Shared notes',
    }),
  ).toEqual({
    initialized: true,
    content: { type: 'markdown', markdown: '# Shared notes' },
  });
  expect(readOfficeCollaborationMetadata(session)).toEqual({
    protocol: 'a3s.office.collaboration',
    version: 1,
    artifactId: 'notes-1',
    kind: 'markdown',
    initialized: true,
  });

  expect(
    initializeOfficeMarkdownCollaboration(session, {
      type: 'markdown',
      markdown: '# Must not overwrite',
    }),
  ).toEqual({
    initialized: false,
    content: { type: 'markdown', markdown: '# Shared notes' },
  });
});

test('requires an explicit synchronized bootstrap before binding', () => {
  const session = createOfficeCollaborationSession({
    artifactId: 'notes-uninitialized',
    kind: 'markdown',
  });

  expect(() => createOfficeMarkdownCollaborationBinding(session)).toThrow(
    /has not been initialized/,
  );
  expect(readOfficeCollaborationMetadata(session)).toBeNull();
});

test('leaves no partial metadata when bootstrap validation fails', () => {
  const session = createOfficeCollaborationSession({
    artifactId: 'notes-invalid-bootstrap',
    kind: 'markdown',
  });
  session.document
    .getText(session.rootName('markdown.source'))
    .insert(0, 'Unattributed source');

  expect(() =>
    initializeOfficeMarkdownCollaboration(session, {
      type: 'markdown',
      markdown: 'Seed',
    }),
  ).toThrow(/contains data without initialized metadata/);
  expect(session.document.getMap(session.rootName('metadata')).size).toBe(0);
  expect(
    session.document.getArray(session.rootName('bootstrap.initializers'))
      .length,
  ).toBe(0);
});

test('detects concurrent bootstrap instead of choosing one initial value', () => {
  const firstDocument = new Y.Doc();
  const secondDocument = new Y.Doc();
  const first = createOfficeCollaborationSession({
    artifactId: 'notes-concurrent-bootstrap',
    document: firstDocument,
    kind: 'markdown',
  });
  const second = createOfficeCollaborationSession({
    artifactId: 'notes-concurrent-bootstrap',
    document: secondDocument,
    kind: 'markdown',
  });
  initializeOfficeMarkdownCollaboration(first, {
    type: 'markdown',
    markdown: 'First seed',
  });
  initializeOfficeMarkdownCollaboration(second, {
    type: 'markdown',
    markdown: 'Second seed',
  });

  exchangeUpdates(firstDocument, secondDocument);

  expect(() => readOfficeCollaborationMetadata(first)).toThrow(
    /Multiple clients initialized/,
  );
  expect(() => readOfficeMarkdownCollaboration(second)).toThrow(
    /Multiple clients initialized/,
  );
});

test('rejects canonical changes outside edit mode and marks agent origins', () => {
  const document = new Y.Doc();
  const writable = createOfficeCollaborationSession({
    actor: { id: 'agent-1', kind: 'agent', name: 'Coding agent' },
    artifactId: 'notes-permissions',
    document,
    kind: 'markdown',
  });
  initializeOfficeMarkdownCollaboration(writable, {
    type: 'markdown',
    markdown: 'Authorized content',
  });
  expect(writable.localOrigin).toMatchObject({
    actorId: 'agent-1',
    kind: 'agent',
  });

  const readOnly = createOfficeCollaborationSession({
    artifactId: 'notes-permissions',
    document,
    kind: 'markdown',
    mode: 'view',
  });
  const binding = createOfficeMarkdownCollaborationBinding(readOnly);
  expect(() => binding.replace('Unauthorized content')).toThrow(
    /cannot modify canonical content/,
  );
  expect(() => readOnly.transact(() => undefined)).toThrow(
    /cannot modify canonical content/,
  );
  expect(readOfficeMarkdownCollaboration(readOnly).markdown).toBe(
    'Authorized content',
  );
});

test('converges concurrent Markdown edits without replacing the shared document', () => {
  const firstDocument = new Y.Doc();
  const first = createOfficeCollaborationSession({
    artifactId: 'notes-2',
    document: firstDocument,
    kind: 'markdown',
  });
  initializeOfficeMarkdownCollaboration(first, {
    type: 'markdown',
    markdown: 'Base',
  });

  const secondDocument = cloneDocument(firstDocument);
  const second = createOfficeCollaborationSession({
    artifactId: 'notes-2',
    document: secondDocument,
    kind: 'markdown',
  });
  const firstBinding = createOfficeMarkdownCollaborationBinding(first);
  const secondBinding = createOfficeMarkdownCollaborationBinding(second);

  firstBinding.replace('Alpha Base');
  secondBinding.replace('Base Omega');
  exchangeUpdates(firstDocument, secondDocument);

  expect(firstBinding.content()).toEqual(secondBinding.content());
  expect(firstBinding.content().markdown).toContain('Alpha');
  expect(firstBinding.content().markdown).toContain('Omega');
  expect(readOfficeCollaborationMetadata(first)?.artifactId).toBe('notes-2');
});

test('tracks only local Markdown operations in each undo manager', () => {
  const firstDocument = new Y.Doc();
  const first = createOfficeCollaborationSession({
    actor: { id: 'ada', name: 'Ada' },
    artifactId: 'notes-3',
    document: firstDocument,
    kind: 'markdown',
  });
  initializeOfficeMarkdownCollaboration(first, {
    type: 'markdown',
    markdown: 'Shared',
  });
  const secondDocument = cloneDocument(firstDocument);
  const second = createOfficeCollaborationSession({
    actor: { id: 'grace', name: 'Grace' },
    artifactId: 'notes-3',
    document: secondDocument,
    kind: 'markdown',
  });
  const firstBinding = createOfficeMarkdownCollaborationBinding(first);
  const secondBinding = createOfficeMarkdownCollaborationBinding(second);

  firstBinding.replace('Shared by Ada');
  exchangeUpdates(firstDocument, secondDocument);
  expect(firstBinding.canUndo()).toBe(true);
  expect(secondBinding.canUndo()).toBe(false);

  secondBinding.replace('Shared by Ada and Grace');
  exchangeUpdates(firstDocument, secondDocument);
  expect(firstBinding.canUndo()).toBe(true);
  expect(secondBinding.canUndo()).toBe(true);

  expect(secondBinding.undo()).toBe(true);
  exchangeUpdates(firstDocument, secondDocument);
  expect(readOfficeMarkdownCollaboration(first).markdown).toBe('Shared by Ada');
  expect(firstBinding.canUndo()).toBe(true);
});

test('reports locality relative to one binding origin', () => {
  const session = createOfficeCollaborationSession({
    artifactId: 'notes-locality',
    kind: 'markdown',
  });
  initializeOfficeMarkdownCollaboration(session, {
    type: 'markdown',
    markdown: 'Shared',
  });
  const first = createOfficeMarkdownCollaborationBinding(session);
  const second = createOfficeMarkdownCollaborationBinding(session);
  const firstLocality: boolean[] = [];
  const secondLocality: boolean[] = [];
  first.subscribe((change) => firstLocality.push(change.local));
  second.subscribe((change) => secondLocality.push(change.local));

  first.replace('Shared by first');

  expect(firstLocality).toEqual([true]);
  expect(secondLocality).toEqual([false]);
  expect(first.canUndo()).toBe(true);
  expect(second.canUndo()).toBe(false);
});

test('does not destroy a host-owned Y.Doc with the session', () => {
  const document = new Y.Doc();
  const session = createOfficeCollaborationSession({
    artifactId: 'notes-4',
    document,
    kind: 'markdown',
  });
  initializeOfficeMarkdownCollaboration(session, {
    type: 'markdown',
    markdown: 'Host owned',
  });
  session.destroy();

  expect(document.getText('a3s.office.markdown.source').toString()).toBe(
    'Host owned',
  );
  document.getMap('still-alive').set('value', true);
  expect(document.getMap('still-alive').get('value')).toBe(true);
});

test('rejects a shared document attached with another identity or kind', () => {
  const document = new Y.Doc();
  const session = createOfficeCollaborationSession({
    artifactId: 'notes-5',
    document,
    kind: 'markdown',
  });
  initializeOfficeMarkdownCollaboration(session, {
    type: 'markdown',
    markdown: 'Identity bound',
  });

  expect(() =>
    createOfficeCollaborationSession({
      artifactId: 'notes-6',
      document,
      kind: 'markdown',
    }),
  ).toThrowError(OfficeCollaborationError);
  expect(() =>
    createOfficeCollaborationSession({
      artifactId: 'notes-5',
      document,
      kind: 'document',
    }),
  ).toThrow(/contains 'markdown' content/);
});

test('preserves UTF-16 pairs when applying a bounded Markdown replacement', () => {
  const session = createOfficeCollaborationSession({
    artifactId: 'notes-emoji',
    kind: 'markdown',
  });
  initializeOfficeMarkdownCollaboration(session, {
    type: 'markdown',
    markdown: 'A😀B',
  });
  const binding = createOfficeMarkdownCollaborationBinding(session);

  binding.replace('A😃B');

  expect(binding.content().markdown).toBe('A😃B');
});

function cloneDocument(source: Y.Doc): Y.Doc {
  const clone = new Y.Doc();
  Y.applyUpdate(clone, Y.encodeStateAsUpdate(source));
  return clone;
}

function exchangeUpdates(first: Y.Doc, second: Y.Doc): void {
  const firstUpdate = Y.encodeStateAsUpdate(first, Y.encodeStateVector(second));
  const secondUpdate = Y.encodeStateAsUpdate(
    second,
    Y.encodeStateVector(first),
  );
  Y.applyUpdate(first, secondUpdate, 'test-network');
  Y.applyUpdate(second, firstUpdate, 'test-network');
}
