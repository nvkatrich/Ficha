# External API references used

Bitrix24 official documentation: https://apidocs.bitrix24.com/api-reference/crm/timeline/comments/crm-timeline-comment-list.html

The `crm.timeline.comment.list` method lists comments for a CRM object using `filter: {ENTITY_ID: deal_id, ENTITY_TYPE: "deal"}`. Its response contains `FILES`; each attachment includes `id`, `name`, `urlDownload`, and the comment includes `ID`, `CREATED`, `COMMENT`. The API returns up to 50 comments per page using `start` pagination. Required scope is `crm` and the current user needs read access to the CRM entity.

Bitrix24 comment overview: https://apidocs.bitrix24.com/api-reference/crm/timeline/comments/index.html

Comment attachments are returned by comment list/get methods. The implementation uses the signed `urlDownload` address, validates HTTPS and the same portal host, downloads only for temporary parsing, and deletes the temporary file afterward. Supported local parsers are XLSX/XLSM/XLS/CSV, DOCX, PPTX, and text-based PDF.
