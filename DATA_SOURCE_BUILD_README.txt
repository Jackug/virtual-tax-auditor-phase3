VTA PHASE 3 - DATA SOURCE APPROVAL MODULE

This build adds controlled Data Source registration, validation, approval, rejection, versioning and review history.

Files to replace/add in the existing Phase 3 service:
1. app.py - replace the existing app.py with this version.
2. templates/data_sources.html - replace existing template.
3. templates/data_source_versions.html - add this template.

No database reset is required. The application creates the new tables automatically at startup.
Existing Knowledge Base tables/data are retained.

Workflow:
Upload -> Validate -> Draft/Pending Review -> Human Approve/Reject/Needs Revision -> Approved source available for downstream use.

Important: this build registers and controls the source; it does NOT yet load approved records into Taxpayer 360. That should be the next controlled step after a real source is uploaded and its columns are mapped.

Accepted data files: CSV, XLSX, XLS.
