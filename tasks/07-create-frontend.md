# Task 07: Create Frontend (Login + Dashboard)

## Objective
Build a vanilla HTML/CSS/JS frontend with a login page and a dashboard for file uploads, metadata input, and individual report downloads. Uses UI UX Pro Max skill for design system.

## Design System (Generated)

**Pattern:** Enterprise Gateway
**Style:** Data-Dense Dashboard
**Colors:**
- Primary: `#0369A1` (security blue)
- Secondary: `#0EA5E9`
- CTA: `#22C55E` (protected green)
- Background: `#F0F9FF`
- Text: `#0C4A6E`

**Typography:** Fira Code / Fira Sans
- Google Fonts: `https://fonts.googleapis.com/css2?family=Fira+Code:wght@400;500;600;700&family=Fira+Sans:wght@300;400;500;600;700`

**Effects:** Hover tooltips, row highlighting on hover, smooth filter animations, data loading spinners

**Anti-patterns to Avoid:**
- No emojis as icons (use SVG: Heroicons/Lucide)
- No ornate design
- No missing filtering
- Add `cursor-pointer` to all clickable elements
- Smooth transitions (150-300ms)
- Light mode text contrast 4.5:1 minimum

## Files to Create
- `static/index.html` — Login page
- `static/dashboard.html` — Main dashboard
- `static/css/style.css` — Styles
- `static/js/app.js` — API calls and DOM logic

## Login Page (`index.html`)

### Layout
- Centered card with login form
- Username + password fields
- "Login" button
- Error message display
- Enterprise Gateway style with security blue theme

### Logic
1. On form submit → `POST /api/login` with credentials
2. Store `access_token` in `localStorage`
3. Redirect to `/` (dashboard) on success
4. Show error on failure

## Dashboard Page (`dashboard.html`)

### Layout (Single Page, Scrollable Sections)

#### Header
- App title "VA/CA Report Automation"
- User info + Logout button
- Floating navbar with proper spacing (`top-4 left-4 right-4`)

#### Section A: File Upload
- Drag-and-drop zone + file picker button
- Accepts `.xlsx` and `.csv` files
- Shows list of selected files with remove buttons
- "Merge Files" button → `POST /api/merge-csv`
- Download merged file on success

#### Section B: Report Metadata Form
Grouped into fieldsets with proper spacing:

**Report Info**
- Report Type: Dropdown (First / Final)
- Report Number: Text input (1.0, 1.1, etc.)
- Client Name: Text input
- Client Short Name: Text input
- Device Type: Text input
- Scope: Dropdown (Server / Firewall / etc.)
- Phase: Text input (First / Final)

**Personnel**
- Prepared By: Text input
- Reviewed By: Text input
- Senior: Dropdown (Vinit / Abhishek / Sravan / Chirag)
- Spokesperson Name: Text input
- Designation: Text input
- Email: Text input

**Dates**
- Assessment Start Date: Date picker
- Assessment Finish Date: Date picker
- Final Retesting Start: Date picker (enabled only when Report Type = Final)
- Final Retesting Finish: Date picker (enabled only when Report Type = Final)
- Released Date: Date picker

**Document Info**
- Document Title: Auto-filled based on Report Type (First Audit Report / Final Audit Report)
- Approved By: Text input (default: "Default")

#### Section C: Generate Buttons
- "Generate Excel Reports" button → `POST /api/report` → returns JSON with download URLs
- "Generate Word Report" button → `POST /api/word` → returns JSON with download URL
- Loading spinners during generation
- Success/error notifications

#### Section D: Download Area (appears after generation)
- List of individual download buttons for each generated file
- Each button triggers `GET /api/download/{session_id}/{file_type}`
- Files download individually (not as ZIP)

### Logic (`app.js`)

```javascript
// Auth
function getToken() → localStorage.getItem("token")
function isLoggedIn() → !!getToken()
function logout() → localStorage.removeItem("token"), redirect to /login

// File Upload
function handleFileSelect(files) → add to file list
function removeFile(index) → remove from file list
async function mergeFiles() → POST /api/merge-csv with FormData

// Report Generation
async function generateReports() → POST /api/report with FormData + metadata
    → On success: show download buttons for each file
async function generateWord() → POST /api/word with FormData + metadata
    → On success: show download button for Word file

// Individual Downloads
async function downloadFile(sessionId, fileType) → GET /api/download/{sessionId}/{fileType}
    → Triggers browser download

// Form Logic
function handleReportTypeChange() → toggle Final Retesting date fields
function updateDocumentTitle() → set title based on Report Type
function buildFormData() → collect all form fields into FormData

// Utility
function showNotification(message, type) → success/error toast
function setLoading(button, loading) → toggle spinner
```

## CSS (`style.css`)

### Design Tokens (from UI UX Pro Max)
```css
:root {
    --color-primary: #0369A1;
    --color-secondary: #0EA5E9;
    --color-cta: #22C55E;
    --color-bg: #F0F9FF;
    --color-text: #0C4A6E;
    --color-surface: #FFFFFF;
    --color-border: #E2E8F0;
    --font-heading: 'Fira Code', monospace;
    --font-body: 'Fira Sans', sans-serif;
    --transition: 150ms ease-in-out;
}
```

### Layout
- Floating navbar with `top-4 left-4 right-4` spacing
- Content padding to account for fixed navbar
- Consistent `max-w-6xl` container width
- Responsive: 375px, 768px, 1024px, 1440px

### Components
- Card-based sections with subtle shadows
- Form inputs with proper focus states
- Buttons with `cursor-pointer` and hover feedback
- Loading spinners (CSS animation)
- Toast notifications (slide in from top-right)

### Interactions
- All clickable elements have `cursor-pointer`
- Hover states with smooth transitions (150-300ms)
- Row highlighting on hover for tables/lists
- Focus states visible for keyboard navigation

## API Integration

All API calls include:
```javascript
headers: {
    "Authorization": `Bearer ${localStorage.getItem("token")}`
}
```

File uploads use `FormData`:
```javascript
const formData = new FormData();
formData.append("file", fileInput.files[0]);
// ... append other fields
const response = await fetch("/api/report", {
    method: "POST",
    headers: { "Authorization": `Bearer ${token}` },
    body: formData
});
```

Individual file downloads:
```javascript
async function downloadFile(sessionId, fileType) {
    const response = await fetch(`/api/download/${sessionId}/${fileType}`, {
        headers: { "Authorization": `Bearer ${token}` }
    });
    const blob = await response.blob();
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${fileType}.xlsx`;
    a.click();
}
```

## Pre-Delivery Checklist (from UI UX Pro Max)

### Visual Quality
- [ ] No emojis used as icons (use SVG: Heroicons/Lucide)
- [ ] All icons from consistent icon set
- [ ] Hover states don't cause layout shift
- [ ] Use theme colors directly (bg-primary) not var() wrapper

### Interaction
- [ ] All clickable elements have `cursor-pointer`
- [ ] Hover states provide clear visual feedback
- [ ] Transitions are smooth (150-300ms)
- [ ] Focus states visible for keyboard navigation

### Light/Dark Mode
- [ ] Light mode text has sufficient contrast (4.5:1 minimum)
- [ ] Glass/transparent elements visible in light mode
- [ ] Borders visible in both modes

### Layout
- [ ] Floating elements have proper spacing from edges
- [ ] No content hidden behind fixed navbars
- [ ] Responsive at 375px, 768px, 1024px, 1440px
- [ ] No horizontal scroll on mobile

### Accessibility
- [ ] All images have alt text
- [ ] Form inputs have labels
- [ ] Color is not the only indicator
- [ ] `prefers-reduced-motion` respected

## Acceptance Criteria
- [ ] Login page loads at `/login` with enterprise design
- [ ] Login stores JWT and redirects to dashboard
- [ ] Dashboard shows file upload area with drag-and-drop
- [ ] Multiple files can be selected and merged
- [ ] Metadata form has all required fields with proper grouping
- [ ] Final Retesting dates enabled only when Report Type = Final
- [ ] Generate Excel button returns JSON with download URLs
- [ ] Generate Word button returns JSON with download URL
- [ ] Individual download buttons appear after generation
- [ ] Each file downloads individually (not bundled)
- [ ] Loading spinners show during generation
- [ ] Error messages displayed on failure
- [ ] JWT included in all API requests
- [ ] Logout clears token and redirects to login
- [ ] Design matches UI UX Pro Max recommendations
- [ ] Responsive on all screen sizes
