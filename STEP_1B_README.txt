VTA STEP 1B - KNOWLEDGE BASE DOCUMENT REGISTRATION, UPLOAD, METADATA AND VERSION CONTROL

This build extends the existing Phase 3 VTA. It does NOT create a new repository/service and does NOT reset the existing database.

Added:
1. knowledge_documents table
2. knowledge_document_versions table
3. Knowledge Base registration form
4. Controlled file upload to uploads/knowledge_base/
5. Metadata capture
6. Draft status pending the Step 1C approval workflow
7. Version history
8. New-version upload without deleting the previous version
9. Download current version
10. Audit events for registration/version upload

Supported file types: PDF, DOC, DOCX, TXT, CSV, XLS, XLSX.
The existing app MAX_CONTENT_LENGTH remains the controlling upload limit.

Next step ONLY after this is tested successfully: Step 1C - human approval/version governance.
