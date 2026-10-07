VTA PHASE 3 — STEP 1C COMPLETE TEMPLATES

This is the complete revised templates folder for Step 1C.

Includes:
- Dashboard approval task for pending Knowledge Base drafts
- Knowledge Approval Queue
- Knowledge Base document registration/upload
- Knowledge Base version history
- Knowledge governance review history
- Existing Phase 3 closed-loop workflow templates

Total templates: 20

Deployment:
1. Replace the existing templates/ folder contents with this templates/ folder.
2. Do NOT replace the database.
3. Do NOT change app.py for this templates-only update.
4. Redeploy/restart the existing Render service.

Required backend routes already present in the Step 1C app.py:
- /knowledge-base/approvals
- /knowledge-base/<document_id>/versions
- /knowledge-base/<document_id>/reviews
- /knowledge-base/<document_id>/review
