# PII Redaction Engine

A robust, enterprise-grade PII detection and redaction engine for Microsoft Word (`.docx`) documents and text.

## Supported PII Entity Types
1. **Full Names (`PERSON`)**: Contextual role-based discovery (Chairman, MD, Promoter, Secretary, etc.), spaCy NER fallback, and two-pass cross-document consistency propagation.
2. **Emails (`EMAIL`)**: RFC-compliant email detection with localized synthetic replacements.
3. **Phone Numbers (`PHONE`)**: Multi-format phone number detection with strict digit length bounds to prevent matching arbitrary financial metrics.
4. **Company Names (`COMPANY`)**: Corporate suffix patterns (`Pvt Ltd`, `LLP`, `Inc`, etc.) with regulatory exclusion filters (`SEBI`, `BSE`, `NSE`, `RBI`, `MCA`, etc.).
5. **Physical Addresses (`ADDRESS`)**: Contextual address prefixes and PIN/postal code pattern detection.
6. **SSNs (`SSN`)**: Standard 9-digit hyphenated SSN patterns.
7. **Credit Cards (`CREDIT_CARD`)**: Regex pattern matching validated by the **Luhn Checksum Algorithm** to prevent false positives on numerical tables.
8. **Dates of Birth (`DATE_OF_BIRTH`)**: Context-aware DOB detection (`DOB:`, `Born on:`, `Date of Birth:`) ensuring standard business/financial dates are preserved.
9. **IP Addresses (`IP_ADDRESS`)**: Strict IPv4 address validation.

## Architectural Highlights
- **Exact Span Replacement & Priority Conflict Resolution**: Replaces tokens by character offset from right-to-left without destructive `str.replace` side-effects.
- **Two-Pass Propagation**: Discovers high-confidence entities in pass 1 and replaces even unlabelled mentions consistently in pass 2.
- **Full DOCX Element Traversal**: Redacts paragraphs, tables, nested tables, headers, and footers.
- **Deterministic Synthetic Replacement**: Uses seeded `Faker("en_IN")` with memoization so identical real entities always map to the same synthetic entity across the document.