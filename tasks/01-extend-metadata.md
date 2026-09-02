# Task 01: Extend EngagementMetadata with New Fields

## Objective
Add all new metadata fields required for the Word document sections (Document Details, Change History, Distribution, Audit Team, Testing Details).

## Files to Modify
- `src/va_ca_automation/metadata/engagement_metadata.py`

## New Fields to Add to `EngagementMetadata` Dataclass

```python
# Report Type & Version
report_type: str = "First"           # "First" or "Final"
report_number: str = "1.0"           # e.g. "1.0", "1.1"

# Client Details
client_short_name: str = ""          # Short name for Document ID (e.g., "VSL")

# Dates
assessment_start_date: str = ""      # YYYY-MM-DD format
assessment_finish_date: str = ""     # YYYY-MM-DD format
final_retesting_start: str = ""      # YYYY-MM-DD (Final only)
final_retesting_finish: str = ""     # YYYY-MM-DD (Final only)
released_date: str = ""              # YYYY-MM-DD

# Document Distribution
spokesperson_name: str = ""
spokesperson_designation: str = ""
spokesperson_email: str = ""

# Audit Team
senior_name: str = ""                # Dropdown: Vinit, Abhishek, Sravan, Chirag
approved_by: str = "Default"         # Default approved by

# Device & Scope
device_type: str = ""                # Added if not already present (check: default_device_type exists)
```

## Implementation Steps

1. Open `src/va_ca_automation/metadata/engagement_metadata.py`
2. Add the new fields after the existing fields in the `EngagementMetadata` dataclass
3. Keep existing fields and their defaults unchanged
4. Add a computed property for document version based on report_type:
   ```python
   @property
   def document_version(self) -> str:
       if self.report_type == "First":
           return "1.0"
       return self.report_number
   ```
5. Add a computed property for document title:
   ```python
   @property
   def document_title(self) -> str:
       if self.report_type == "First":
           return "First Audit Report"
       return "Final Audit Report"
   ```
6. Add a property to get assessment date range:
   ```python
   @property
   def assessment_date_range(self) -> str:
       return f"{self.assessment_start_date} to {self.assessment_finish_date}"
   ```
7. Add a property for testing dates (conditional on report_type):
   ```python
   @property
   def first_audit_dates(self) -> str:
       if self.report_type == "First":
           return f"{self.assessment_start_date} to {self.assessment_finish_date}"
       return "NA"

   @property
   def final_retesting_dates(self) -> str:
       if self.report_type == "Final":
           return f"{self.final_retesting_start} to {self.final_retesting_finish}"
       return "Revalidation not performed"
   ```

## Acceptance Criteria
- [ ] All new fields added with correct defaults
- [ ] Existing fields unchanged
- [ ] Computed properties work correctly for both "First" and "Final" report types
- [ ] No import errors
