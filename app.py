"""
Azure DevOps Test Case Uploader 
A Python Shiny application for processing and uploading test cases to Azure DevOps

Author: Biostatistics & Data Science
Version: 2.0.5
"""

from shiny import App, ui, render, reactive
import pandas as pd
import requests
import json
import base64
import time
import re
from datetime import datetime

# Application version
VERSION = "2.0.9"

# Configuration
BATCH_SIZE = 1000  # Maximum test cases per batch

# ============================================================================
# USER INTERFACE
# ============================================================================

app_ui = ui.page_fluid(
    ui.tags.head(
        ui.tags.style("""
            .success-badge { background-color: #28a745; color: white; padding: 2px 8px; border-radius: 3px; }
            .error-badge { background-color: #dc3545; color: white; padding: 2px 8px; border-radius: 3px; }
            .progress-container { margin: 20px 0; }
            .info-box { background-color: #d4edda; padding: 10px; border-radius: 5px; color: #155724; margin: 10px 0; }
        """)
    ),
    
    ui.panel_title("🚀 Azure DevOps Test Case Uploader"),
    
    ui.div(
        ui.tags.span(f"Version {VERSION}", style="color: #6c757d; font-size: 0.9em; margin-left: 10px;"),
        style="margin-bottom: 15px;"
    ),
    
    ui.layout_sidebar(
        ui.sidebar(
            ui.h4("Azure DevOps Configuration"),
            ui.input_text("organization", "Organization Name", 
                         placeholder="e.g., mycompany"),
            ui.input_text("project", "Project Name", 
                         placeholder="e.g., MyProject"),
            ui.input_password("pat_token", "Personal Access Token (PAT)", 
                            placeholder="Enter your Azure DevOps PAT"),
            ui.hr(),
            
            ui.h4("File Upload"),
            ui.input_file("csv_file", "Upload Test Cases CSV", 
                         accept=[".csv"], multiple=False),
            ui.hr(),
            
            ui.h4("Optional Overrides"),
            ui.input_text("area_path", "Area Path", 
                         placeholder="Leave blank to use from CSV"),
            
            ui.input_checkbox("dry_run", "Dry Run (Preview only)", value=False),
            ui.hr(),
            
            ui.input_action_button("process_btn", "1. Process CSV", 
                                  class_="btn-primary w-100 mb-2"),
            ui.download_button("download_processed", "📥 Download Processed CSV", 
                              class_="btn-secondary w-100 mb-2"),
            ui.input_action_button("validate_btn", "2. Validate Connection", 
                                  class_="btn-info w-100 mb-2"),
            ui.input_action_button("upload_btn", "3. Upload to Azure DevOps", 
                                  class_="btn-success w-100"),
            width=380
        ),
        
        ui.navset_tab(
            ui.nav_panel("📊 Data Summary",
                ui.card(
                    ui.card_header("Upload Statistics"),
                    ui.output_ui("data_statistics")
                ),
                ui.card(
                    ui.card_header("Data Preview (First 10 rows)"),
                    ui.output_table("data_preview")
                )
            ),
            
            ui.nav_panel("📋 Test Cases",
                ui.card(
                    ui.card_header("Processed Test Cases Summary"),
                    ui.output_ui("test_cases_summary")
                ),
                ui.card(
                    ui.card_header("Test Case Examples"),
                    ui.output_ui("test_cases_display")
                )
            ),
            
            ui.nav_panel("📤 Upload Progress",
                ui.card(
                    ui.card_header("Upload Status"),
                    ui.output_ui("upload_progress")
                ),
                ui.card(
                    ui.card_header("Detailed Results"),
                    ui.output_table("upload_results")
                )
            ),
            
            ui.nav_panel("⚙️ Settings & Help",
                ui.card(
                    ui.card_header("Application Info"),
                    ui.markdown(f"""
                    **Version:** {VERSION}  
                    **Author:** Biostatistics & Data Science  
                    **Purpose:** Process and upload test cases to Azure DevOps
                    """)
                ),
                ui.card(
                    ui.card_header("Connection Status"),
                    ui.output_text_verbatim("connection_status")
                ),
                ui.card(
                    ui.card_header("Quick Help"),
                    ui.markdown("""
                    **Steps to Use:**
                    1. Enter your Azure DevOps organization and project name
                    2. Paste your Personal Access Token (PAT)
                    3. Upload your CSV file
                    4. Click "Process CSV" to structure test cases
                    5. Click "Download Processed CSV" to review (optional)
                    6. Click "Validate Connection" to test Azure DevOps access
                    7. Click "Upload to Azure DevOps" to create test cases
                    
                    **Tips:**
                    - Use "Dry Run" to preview without creating work items
                    - Check Data Summary to verify CSV loaded correctly
                    - Download processed CSV to review structure before upload
                    - Monitor Upload Progress for real-time status
                    
                    **Required CSV Columns:**
                    - `Custom.TestCaseClassification` (Form Level, Field Level, Edit Check Level)
                    - `Custom.FormName`
                    - `Custom.FieldorEditCheckText`
                    - `Custom.TestingTier` (Tier 1, Tier 2, or Tier 3)
                    
                    **Optional CSV Columns:**
                    - `Custom.FieldName` (for Field Level)
                    - `Custom.EditCheckName` (for Edit Check Level)
                    - `Area Path`
                    - `State`
                    
                    **Testing Tier Specifications:**
                    - **Field Level - Tier 1**: Validates variable name, label, SAS label, format, length, pick lists, value choices, calculation
                    - **Field Level - Tier 2**: Validates variable format, length, pick lists, value choices, calculation
                    - **Field Level - Tier 3**: Validates variable format, length, calculation
                    - **Edit Check Level - Tier 1**: Tests positive/negative cases, range checks, query text accuracy, and programming sufficiency
                    - **Edit Check Level - Tier 2**: Tests query text accuracy and programming sufficiency
                    - **Edit Check Level - Tier 3**: Tests programming sufficiency only
                    - **Form Level**: Testing Tier is captured but does not affect test case generation
                    """)
                )
            )
        )
    )
)


# ============================================================================
# SERVER LOGIC
# ============================================================================

def server(input, output, session):
    # Reactive values to store data
    uploaded_data = reactive.Value(None)
    processed_test_cases = reactive.Value(None)
    upload_results_data = reactive.Value(None)
    validation_status = reactive.Value(None)
    upload_progress_info = reactive.Value(None)
    
    def create_auth_header(pat_token):
        """Create authorization header for Azure DevOps API"""
        auth_string = f":{pat_token}"
        auth_bytes = auth_string.encode('ascii')
        auth_b64 = base64.b64encode(auth_bytes).decode('ascii')
        return {
            'Authorization': f'Basic {auth_b64}',
            'Content-Type': 'application/json-patch+json'
        }
    
    def escape_xml(text):
        """Escape special characters for XML"""
        if not text:
            return ""
        return str(text).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;').replace("'", '&apos;')
    
    def strip_html_tags(text):
        """Remove HTML tags from text"""
        if not text or pd.isna(text):
            return ""
        # Convert to string
        text = str(text)
        # Remove HTML tags using regex
        clean_text = re.sub(r'<[^>]+>', '', text)
        # Decode common HTML entities
        clean_text = clean_text.replace('&nbsp;', ' ')
        clean_text = clean_text.replace('&amp;', '&')
        clean_text = clean_text.replace('&lt;', '<')
        clean_text = clean_text.replace('&gt;', '>')
        clean_text = clean_text.replace('&quot;', '"')
        clean_text = clean_text.replace('&#39;', "'")
        # Remove extra whitespace
        clean_text = ' '.join(clean_text.split())
        return clean_text.strip()
    
    # ========================================================================
    # DATA LOADING
    # ========================================================================
    
    @reactive.Effect
    @reactive.event(input.csv_file)
    def load_csv():
        """Load and validate CSV file"""
        file_info = input.csv_file()
        if file_info is None or len(file_info) == 0:
            uploaded_data.set(None)
            return
        
        try:
            file_path = file_info[0]["datapath"]
            df = pd.read_csv(file_path)
            df.columns = df.columns.str.strip()
            
            uploaded_data.set(df)
            
        except Exception as e:
            uploaded_data.set(None)
            print(f"Error loading CSV: {e}")
    
    # ========================================================================
    # DATA PREVIEW OUTPUTS
    # ========================================================================
    
    @output
    @render.ui
    def data_statistics():
        """Display data statistics"""
        df = uploaded_data.get()
        if df is None:
            return ui.div(
                ui.tags.p("⚠️ No data loaded. Please upload a CSV file.", 
                         class_="text-warning")
            )
        
        stats = []
        stats.append(ui.tags.p(ui.tags.strong("Total Rows: "), f"{len(df):,}"))
        stats.append(ui.tags.p(ui.tags.strong("Total Columns: "), f"{len(df.columns)}"))
        
        # Check for required columns
        required_cols = ['Custom.TestCaseClassification', 'Custom.FormName']
        missing_cols = [col for col in required_cols if col not in df.columns]
        
        if missing_cols:
            stats.append(ui.tags.hr())
            stats.append(ui.tags.div(
                ui.tags.h5("⚠️ Warning: Missing Required Columns", class_="text-danger"),
                ui.tags.p("The following required columns are missing:"),
                ui.tags.ul(*[ui.tags.li(col) for col in missing_cols]),
                ui.tags.p("Processing will fail. Please check your CSV file."),
                style="background-color: #f8d7da; padding: 10px; border-radius: 5px; color: #721c24;"
            ))
            return ui.div(*stats)
        
        if 'Custom.TestCaseClassification' in df.columns:
            stats.append(ui.tags.hr())
            stats.append(ui.tags.h5("Classification Breakdown:"))
            counts = df['Custom.TestCaseClassification'].value_counts()
            for classification, count in counts.items():
                stats.append(ui.tags.p(
                    f"• {classification}: ",
                    ui.tags.strong(f"{count:,}"),
                    " items"
                ))
        
        if 'Custom.FormName' in df.columns:
            unique_forms = df['Custom.FormName'].nunique()
            stats.append(ui.tags.hr())
            stats.append(ui.tags.p(
                ui.tags.strong("Unique Forms: "),
                f"{unique_forms}"
            ))
        
        if 'Custom.TestingTier' in df.columns:
            # Show Testing Tier breakdown by Classification
            stats.append(ui.tags.hr())
            stats.append(ui.tags.h5("Testing Tier Distribution:"))

            for classification in df['Custom.TestCaseClassification'].dropna().unique():
                class_df = df[df['Custom.TestCaseClassification'] == classification]

                # Count tiers including missing values
                tier_counts = {}
                for tier in class_df['Custom.TestingTier']:
                    if pd.notna(tier) and str(tier).strip():
                        tier_str = str(tier).strip()
                        tier_counts[tier_str] = tier_counts.get(tier_str, 0) + 1
                    else:
                        tier_counts['Missing'] = tier_counts.get('Missing', 0) + 1

                if len(tier_counts) > 0:
                    stats.append(ui.tags.p(ui.tags.strong(f"{classification}:")))

                    # Display in specific order: Tier 1, Tier 2, Tier 3, Missing
                    tier_order = ['Tier 1', 'Tier 2', 'Tier 3', 'Missing']
                    for tier in tier_order:
                        if tier in tier_counts:
                            stats.append(ui.tags.p(
                                f"  • {tier}: ",
                                ui.tags.strong(f"{tier_counts[tier]:,}"),
                                " items"
                            ))

                    # Display any other tiers that weren't in the predefined order
                    for tier, count in tier_counts.items():
                        if tier not in tier_order:
                            stats.append(ui.tags.p(
                                f"  • {tier}: ",
                                ui.tags.strong(f"{count:,}"),
                                " items"
                            ))
        
        return ui.div(*stats)
    
    @output
    @render.table
    def data_preview():
        """Show preview of uploaded data"""
        df = uploaded_data.get()
        if df is None:
            return pd.DataFrame({"Message": ["No data loaded"]})
        # Replace NaN values with empty strings for display
        return df.head(10).fillna('')
    
    # ========================================================================
    # TEST CASE PROCESSING
    # ========================================================================
    
    @reactive.Effect
    @reactive.event(input.process_btn)
    def process_test_cases():
        """Process CSV into test case structure"""
        df = uploaded_data.get()
        if df is None:
            processed_test_cases.set(None)
            return
        
        # Validate required columns exist
        required_cols = ['Custom.TestCaseClassification', 'Custom.FormName']
        missing_cols = [col for col in required_cols if col not in df.columns]
        
        if missing_cols:
            # Set error state
            processed_test_cases.set({
                'error': True,
                'message': f"Missing required columns: {', '.join(missing_cols)}. Please check your CSV file."
            })
            return
        
        test_cases = []
        forms_with_null_tiers = []  # Track forms with null Form level testing tiers

        # Group by Form Name
        unique_forms = df['Custom.FormName'].dropna().unique()

        for form_name in unique_forms:
            form_data = df[df['Custom.FormName'] == form_name]

            # Get form level testing tier if it exists (for inheritance by Field/Edit Check levels)
            form_level_testing_tier = ''
            form_level = form_data[form_data['Custom.TestCaseClassification'] == 'Form Level']
            if len(form_level) > 0:
                first_form_level = form_level.iloc[0]
                tier_value = first_form_level.get('Custom.TestingTier', '')
                if pd.notna(tier_value) and str(tier_value).strip():
                    form_level_testing_tier = str(tier_value).strip()
                else:
                    # Track forms with null/empty Form level testing tiers
                    forms_with_null_tiers.append(form_name)

# Process Form Level items (standalone test cases)
            for _, row in form_level.iterrows():
                # Define standard Form Level test steps
                form_level_steps = [
                    {
                        'step_number': 1,
                        'action': 'Review all questions/fields on CRF against CRF specification',
                        'expected': 'All expected fields are present on eCRF and text is accurate'
                    },
                    {
                        'step_number': 2,
                        'action': 'Review order of CRF questions',
                        'expected': 'Order of eCRF questions is logical and accurate'
                    },
                    {
                        'step_number': 3,
                        'action': 'Review required data on each CRF',
                        'expected': 'All required fields are indicated as such before or after save or proper queries are issued for missing data'
                    },
                    {
                        'step_number': 4,
                        'action': 'Review field dynamics against specification',
                        'expected': 'All eCRF questions/fields are available as expected and hidden / not available as expected.'
                    }
                ]
                
                test_cases.append({
                    'type': 'standalone',
                    'title': f"{form_name} - Form Review",
                    'form_name': form_name,
                    'classification': 'Form Level',
                    'testing_tier': str(row.get('Custom.TestingTier', '')),
                    'description': strip_html_tags(row.get('Custom.FieldorEditCheckText', '')),
                    'area_path': str(row.get('Area Path', '')),
                    'state': str(row.get('State', 'Design')),
                    'steps': form_level_steps
                })
            
            # Process Field Level items (group by Testing Tier)
            field_level = form_data[form_data['Custom.TestCaseClassification'] == 'Field Level']
            if len(field_level) > 0:
                # Create a copy and apply inheritance from form level testing tier
                field_level = field_level.copy()
                if form_level_testing_tier:
                    # Fill empty/missing testing tier values with form level tier
                    field_level['Custom.TestingTier'] = field_level['Custom.TestingTier'].apply(
                        lambda x: form_level_testing_tier if pd.isna(x) or not str(x).strip() else x
                    )

                # Group by Testing Tier (now with inherited values)
                for testing_tier in field_level['Custom.TestingTier'].unique():
                    tier_data = field_level[field_level['Custom.TestingTier'] == testing_tier]
                    if len(tier_data) == 0:
                        continue
                    
                    steps = []
                    for idx, row in tier_data.iterrows():
                        field_name = str(row.get('Custom.FieldName', ''))
                        field_text = strip_html_tags(row.get('Custom.FieldorEditCheckText', ''))
                        tier_value = str(row.get('Custom.TestingTier', '')).strip()
                        
                        # Determine expected result based on Testing Tier
                        if tier_value == 'Tier 1':
                            expected = "Correct variable name, label, SAS label, format, length, pick lists, value choices, calculation."
                        elif tier_value == 'Tier 2':
                            expected = "Correct variable format, length, pick lists, value choices, calculation."
                        elif tier_value == 'Tier 3':
                            expected = "Correct variable format, length, calculation."
                        else:
                            # Default if tier is not specified or unknown
                            expected = f"Field '{field_name}' validates correctly. {field_text}"
                        
                        # Action format: "Review field: FieldName (FieldText)"
                        action = f"Review field: {field_name}"
                        if field_text:
                            action += f" ({field_text})"
                        
                        steps.append({
                            'step_number': len(steps) + 1,
                            'action': action,
                            'expected': expected,
                            'field_name': field_name
                        })
                    
                    first_row = tier_data.iloc[0]
                    test_cases.append({
                        'type': 'field_reviews',
                        'title': f"{form_name} - Field Reviews",
                        'form_name': form_name,
                        'classification': 'Field Level',
                        'testing_tier': str(first_row.get('Custom.TestingTier', '')),
                        'description': f"Field-level validation for form {form_name}. Total fields: {len(steps)}",
                        'area_path': str(first_row.get('Area Path', '')),
                        'state': str(first_row.get('State', 'Design')),
                        'steps': steps
                    })
            
            # Process Edit Check Level items (group by Testing Tier)
            edit_check_level = form_data[form_data['Custom.TestCaseClassification'] == 'Edit Check Level']
            if len(edit_check_level) > 0:
                # Create a copy and apply inheritance from form level testing tier
                edit_check_level = edit_check_level.copy()
                if form_level_testing_tier:
                    # Fill empty/missing testing tier values with form level tier
                    edit_check_level['Custom.TestingTier'] = edit_check_level['Custom.TestingTier'].apply(
                        lambda x: form_level_testing_tier if pd.isna(x) or not str(x).strip() else x
                    )

                # Group by Testing Tier (now with inherited values)
                for testing_tier in edit_check_level['Custom.TestingTier'].unique():
                    tier_data = edit_check_level[edit_check_level['Custom.TestingTier'] == testing_tier]
                    if len(tier_data) == 0:
                        continue
                    
                    steps = []
                    for idx, row in tier_data.iterrows():
                        edit_check_name = str(row.get('Custom.EditCheckName', ''))
                        edit_check_text = strip_html_tags(row.get('Custom.FieldorEditCheckText', ''))
                        tier_value = str(row.get('Custom.TestingTier', '')).strip()
                        
                        # Determine expected result based on Testing Tier
                        if tier_value == 'Tier 1':
                            expected = "Positive Test in System Passes, Negative Test in System passes, Range Checks (high / low) pass, Query Text is correct and accurate and not leading. Sufficient Edit Checks are programmed"
                        elif tier_value == 'Tier 2':
                            expected = "Query Text is correct and accurate and not leading. Sufficient Edit Checks are programmed"
                        elif tier_value == 'Tier 3':
                            expected = "Sufficient Edit Checks are programmed."
                        else:
                            # Default if tier is not specified or unknown
                            expected = f"Edit check '{edit_check_name}' functions correctly. {edit_check_text}"
                        
                        # Action format: "Review Edit Check: EditCheckName (EditCheckText)"
                        action = f"Review Edit Check: {edit_check_name}"
                        if edit_check_text:
                            action += f" ({edit_check_text})"
                        
                        steps.append({
                            'step_number': len(steps) + 1,
                            'action': action,
                            'expected': expected,
                            'edit_check_name': edit_check_name
                        })
                    
                    first_row = tier_data.iloc[0]
                    test_cases.append({
                        'type': 'edit_check_reviews',
                        'title': f"{form_name} - Edit Check Reviews",
                        'form_name': form_name,
                        'classification': 'Edit Check Level',
                        'testing_tier': str(first_row.get('Custom.TestingTier', '')),
                        'description': f"Edit check validation for form {form_name}. Total checks: {len(steps)}",
                        'area_path': str(first_row.get('Area Path', '')),
                        'state': str(first_row.get('State', 'Design')),
                        'steps': steps
                    })

        # Store test cases along with metadata about null Form level testing tiers
        processed_data = {
            'test_cases': test_cases,
            'forms_with_null_tiers': forms_with_null_tiers
        }
        processed_test_cases.set(processed_data)
    
    # ========================================================================
    # TEST CASE DISPLAY
    # ========================================================================
    
    @output
    @render.ui
    def test_cases_summary():
        """Show summary of processed test cases"""
        processed_data = processed_test_cases.get()
        if processed_data is None:
            return ui.div(
                ui.tags.p("⚠️ No test cases processed. Click 'Process CSV' first.",
                         class_="text-warning")
            )

        # Check for error state
        if isinstance(processed_data, dict) and processed_data.get('error'):
            return ui.div(
                ui.tags.h4("❌ Processing Error", class_="text-danger"),
                ui.tags.p(processed_data.get('message', 'Unknown error occurred')),
                ui.tags.hr(),
                ui.tags.p("Please ensure your CSV file contains the following required columns:"),
                ui.tags.ul(
                    ui.tags.li("Custom.TestCaseClassification"),
                    ui.tags.li("Custom.FormName"),
                    ui.tags.li("Custom.FieldorEditCheckText")
                )
            )

        # Extract test cases and null tier info from processed data
        test_cases = processed_data.get('test_cases', [])
        forms_with_null_tiers = processed_data.get('forms_with_null_tiers', [])

        type_counts = {}
        total_steps = 0
        for tc in test_cases:
            tc_type = tc['type']
            type_counts[tc_type] = type_counts.get(tc_type, 0) + 1
            total_steps += len(tc['steps'])
        
        summary = [
            ui.tags.h4(f"Total Test Cases: {len(test_cases)}"),
            ui.tags.hr(),
            ui.tags.p(ui.tags.strong("Breakdown by Type:")),
        ]
        
        type_names = {
            'standalone': '📄 Standalone (Form Level)',
            'field_reviews': '🔍 Field Reviews (grouped by tier)',
            'edit_check_reviews': '✅ Edit Check Reviews (grouped by tier)'
        }
        
        for tc_type, count in type_counts.items():
            summary.append(ui.tags.p(
                f"{type_names.get(tc_type, tc_type)}: ",
                ui.tags.strong(f"{count}")
            ))
        
        summary.append(ui.tags.hr())
        summary.append(ui.tags.p(
            ui.tags.strong("Total Steps: "),
            f"{total_steps:,}"
        ))
        
        summary.append(ui.tags.hr())

        # Add warning if there are forms with null Form level testing tiers
        if forms_with_null_tiers:
            summary.append(ui.tags.div(
                ui.tags.h5("⚠️ Warning: Forms with Null Testing Tiers", class_="text-warning"),
                ui.tags.p(f"The following {len(forms_with_null_tiers)} form(s) have null or empty Form Level Testing Tier values:"),
                ui.tags.ul(*[ui.tags.li(form) for form in forms_with_null_tiers]),
                ui.tags.p("Field Level and Edit Check Level items for these forms will not inherit a testing tier value."),
                style="background-color: #fff3cd; padding: 10px; border-radius: 5px; color: #856404; margin-bottom: 15px;"
            ))

        summary.append(ui.tags.div(
            ui.tags.p(
                "✅ Processed file ready! ",
                ui.tags.br(),
                "Click 'Download Processed CSV' to export this structure.",
                ui.tags.br(),
                ui.tags.br(),
                ui.tags.em("Note: Field Level and Edit Check Level test cases are grouped by Testing Tier, with tier-specific expected results.")
            ),
            class_="info-box"
        ))

        return ui.div(*summary)
    
    @output
    @render.ui
    def test_cases_display():
        """Display sample test cases"""
        processed_data = processed_test_cases.get()
        if processed_data is None:
            return ui.p("No test cases to display.")

        # Check for error state
        if isinstance(processed_data, dict) and processed_data.get('error'):
            return ui.div(
                ui.tags.p("Cannot display test cases due to processing error.", class_="text-danger")
            )

        # Extract test cases from processed data
        test_cases = processed_data.get('test_cases', [])
        display_cases = test_cases[:5]
        panels = []
        
        for i, tc in enumerate(display_cases):
            step_info = ""
            if tc['steps']:
                steps_list = [
                    ui.tags.li(
                        ui.tags.strong(f"Step {step['step_number']}: "),
                        step['action'][:80] + ("..." if len(step['action']) > 80 else "")
                    ) for step in tc['steps'][:3]
                ]
                step_info = ui.tags.ul(*steps_list)
                if len(tc['steps']) > 3:
                    step_info = ui.div(
                        step_info,
                        ui.tags.em(f"... and {len(tc['steps']) - 3} more steps")
                    )
            
            content_items = [
                ui.tags.p(ui.tags.strong("Title: "), tc['title']),
                ui.tags.p(ui.tags.strong("Type: "), tc['type'].replace('_', ' ').title()),
                ui.tags.p(ui.tags.strong("Form: "), tc['form_name']),
                ui.tags.p(ui.tags.strong("Classification: "), tc['classification'])
            ]
            
            # Add Testing Tier if it exists
            if tc.get('testing_tier'):
                content_items.append(ui.tags.p(ui.tags.strong("Testing Tier: "), tc['testing_tier']))
            
            content_items.extend([
                ui.tags.p(ui.tags.strong("Total Steps: "), str(len(tc['steps']))),
                ui.tags.hr(),
                step_info if tc['steps'] else ui.tags.em("No steps (standalone test case)")
            ])
            
            content = ui.div(*content_items)
            
            panels.append(ui.accordion_panel(
                f"Example {i+1}: {tc['title']}", 
                content
            ))
        
        if len(test_cases) > 5:
            panels.append(ui.accordion_panel(
                f"➕ {len(test_cases) - 5} more test cases...",
                ui.p("All test cases will be created when you click 'Upload to Azure DevOps'.")
            ))
        
        return ui.accordion(*panels, id="examples_accordion", multiple=False)
    
    # ========================================================================
    # DOWNLOAD PROCESSED CSV
    # ========================================================================
    
    @render.download(filename=lambda: f"processed_test_cases_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")
    def download_processed():
        """Download processed test cases as CSV in Azure DevOps format"""
        processed_data = processed_test_cases.get()

        if processed_data is None:
            # Return empty CSV with message
            df = pd.DataFrame({"Message": ["No processed test cases available. Please process CSV first."]})
            yield df.to_csv(index=False, na_rep='')
            return

        # Check for error state
        if isinstance(processed_data, dict) and processed_data.get('error'):
            # Return CSV with error message
            df = pd.DataFrame({"Error": [processed_data.get('message', 'Unknown error occurred')]})
            yield df.to_csv(index=False, na_rep='')
            return

        # Extract test cases from processed data
        test_cases = processed_data.get('test_cases', [])

        # Convert test cases to Azure DevOps hierarchical format
        # Each test case = 1 header row + N step rows
        # IMPORTANT: Column order must match exactly as specified
        rows = []

        for tc in test_cases:
            # Add test case header row (metadata only, steps empty)
            header_row = {
                'ID': '',  # Will be assigned by Azure DevOps on import
                'Work Item Type': 'Test Case',
                'Title': tc['title'],
                'Test Step': '',
                'Step Action': '',
                'Step Expected': '',
                'Custom.EditCheckName': '',
                'Custom.FieldName': '',
                'Custom.FormName': tc['form_name'],
                'Custom.TestCaseClassification': tc['classification'],
                'Custom.TestingTier': tc.get('testing_tier', ''),
                'Area Path': tc['area_path'] if tc['area_path'] else '',
                'Assigned To': '',
                'State': tc['state']
            }
            rows.append(header_row)
            
            # Add step rows (only step fields filled, rest empty)
            for step in tc['steps']:
                step_row = {
                    'ID': '',
                    'Work Item Type': '',
                    'Title': '',
                    'Test Step': step['step_number'],
                    'Step Action': step['action'],
                    'Step Expected': step['expected'],
                    'Custom.EditCheckName': '',
                    'Custom.FieldName': '',
                    'Custom.FormName': '',
                    'Custom.TestCaseClassification': '',
                    'Custom.TestingTier': '',
                    'Area Path': '',
                    'Assigned To': '',
                    'State': ''
                }
                rows.append(step_row)
        
        # Create DataFrame with exact column order
        columns_order = [
            'ID',
            'Work Item Type',
            'Title',
            'Test Step',
            'Step Action',
            'Step Expected',
            'Custom.EditCheckName',
            'Custom.FieldName',
            'Custom.FormName',
            'Custom.TestCaseClassification',
            'Custom.TestingTier',
            'Area Path',
            'Assigned To',
            'State'
        ]
        
        df = pd.DataFrame(rows, columns=columns_order)
        
        # Export CSV with NaN values as blanks
        yield df.to_csv(index=False, na_rep='')
    
    # ========================================================================
    # CONNECTION VALIDATION
    # ========================================================================
    
    @reactive.Effect
    @reactive.event(input.validate_btn)
    def validate_connection():
        """Validate Azure DevOps connection"""
        organization = input.organization()
        project = input.project()
        pat_token = input.pat_token()
        
        if not all([organization, project, pat_token]):
            validation_status.set({
                'success': False,
                'message': 'Please fill in organization, project, and PAT token.'
            })
            return
        
        try:
            base_url = f"https://dev.azure.com/{organization}/{project}/_apis"
            headers = create_auth_header(pat_token)
            
            # Test connection by getting project info
            test_url = f"{base_url}/wit/workitemtypes?api-version=7.0"
            response = requests.get(test_url, headers=headers, timeout=10)
            
            if response.status_code == 200:
                validation_status.set({
                    'success': True,
                    'message': f'✅ Connection successful! Project: {project}'
                })
            else:
                validation_status.set({
                    'success': False,
                    'message': f'❌ Connection failed: HTTP {response.status_code}\nResponse: {response.text[:200]}'
                })
                
        except Exception as e:
            validation_status.set({
                'success': False,
                'message': f'❌ Error: {str(e)}'
            })
    
    @output
    @render.text
    def connection_status():
        """Display connection validation status"""
        status = validation_status.get()
        if status is None:
            return "Click 'Validate Connection' to test Azure DevOps access."
        return status['message']
    
    # ========================================================================
    # AZURE DEVOPS UPLOAD
    # ========================================================================
    
    @reactive.Effect
    @reactive.event(input.upload_btn)
    async def upload_to_devops():
        """Upload test cases to Azure DevOps"""
        processed_data = processed_test_cases.get()
        organization = input.organization()
        project = input.project()
        pat_token = input.pat_token()
        dry_run = input.dry_run()

        if not all([processed_data, organization, project, pat_token]):
            upload_progress_info.set({
                'status': 'error',
                'message': 'Missing required configuration. Please fill all fields.'
            })
            return

        # Check for error state in processed test cases
        if isinstance(processed_data, dict) and processed_data.get('error'):
            upload_progress_info.set({
                'status': 'error',
                'message': f"Cannot upload: {processed_data.get('message', 'Unknown error in test cases')}"
            })
            return

        # Extract test cases from processed data
        test_cases = processed_data.get('test_cases', [])

        base_url = f"https://dev.azure.com/{organization}/{project}/_apis"
        headers = create_auth_header(pat_token)

        total = len(test_cases)
        all_results = []
        
        # Split into batches if more than BATCH_SIZE
        num_batches = (total + BATCH_SIZE - 1) // BATCH_SIZE  # Ceiling division
        
        if num_batches > 1:
            upload_progress_info.set({
                'status': 'info',
                'message': f'Large upload detected: {total} test cases will be processed in {num_batches} batches of up to {BATCH_SIZE} items each.'
            })
            time.sleep(2)  # Give user time to see the message
        
        for batch_num in range(num_batches):
            start_idx = batch_num * BATCH_SIZE
            end_idx = min(start_idx + BATCH_SIZE, total)
            batch_test_cases = test_cases[start_idx:end_idx]
            batch_size = len(batch_test_cases)
            
            batch_label = f"Batch {batch_num + 1}/{num_batches}" if num_batches > 1 else ""
            
            for idx, tc in enumerate(batch_test_cases, 1):
                overall_idx = start_idx + idx
                
                # Update progress
                progress_msg = f'{batch_label} - Processing {tc["title"]} ({overall_idx}/{total})' if batch_label else f'Processing: {tc["title"]} ({overall_idx}/{total})'
                upload_progress_info.set({
                    'status': 'running',
                    'current': overall_idx,
                    'total': total,
                    'batch': batch_num + 1 if num_batches > 1 else None,
                    'total_batches': num_batches if num_batches > 1 else None,
                    'message': progress_msg
                })
                
                try:
                    if dry_run:
                        # Simulate upload for dry run
                        time.sleep(0.05)
                        all_results.append({
                            'Batch': batch_num + 1 if num_batches > 1 else 'N/A',
                            'Title': tc['title'],
                            'Status': 'Dry Run',
                            'Work Item ID': 'N/A',
                            'Steps': len(tc['steps']),
                            'Timestamp': datetime.now().strftime('%H:%M:%S')
                        })
                        continue
                    
                    # Build work item data
                    area_path = input.area_path() or tc['area_path']
                    
                    work_item_data = [
                        {"op": "add", "path": "/fields/System.Title", "value": tc['title']},
                        {"op": "add", "path": "/fields/System.State", "value": tc['state']},
                        {"op": "add", "path": "/fields/System.Description", "value": tc['description']}
                    ]
                    
                    if area_path:
                        work_item_data.append({"op": "add", "path": "/fields/System.AreaPath", "value": area_path})
                    
                    # Add custom fields (if they exist in your project template)
                    try:
                        work_item_data.append({"op": "add", "path": "/fields/Custom.TestCaseClassification", "value": tc['classification']})
                        work_item_data.append({"op": "add", "path": "/fields/Custom.FormName", "value": tc['form_name']})
                        # Add Testing Tier only if it exists and has a value
                        if tc.get('testing_tier'):
                            work_item_data.append({"op": "add", "path": "/fields/Custom.TestingTier", "value": tc['testing_tier']})
                    except:
                        pass  # Custom fields may not exist in all projects
                    
                    # Create work item
                    create_url = f"{base_url}/wit/workitems/$Test Case?api-version=7.0"
                    response = requests.post(create_url, headers=headers, json=work_item_data, timeout=30)
                    
                    if response.status_code in [200, 201]:
                        work_item = response.json()
                        work_item_id = work_item['id']
                        
                        # Add steps if any
                        if tc['steps']:
                            steps_xml = '<steps id="0" last="' + str(len(tc['steps'])) + '">'
                            for step in tc['steps']:
                                action = escape_xml(step['action'])
                                expected = escape_xml(step['expected'])
                                steps_xml += f'''<step id="{step['step_number']}" type="ValidateStep">
                                    <parameterizedString isformatted="true">&lt;DIV&gt;&lt;P&gt;{action}&lt;/P&gt;&lt;/DIV&gt;</parameterizedString>
                                    <parameterizedString isformatted="true">&lt;DIV&gt;&lt;P&gt;{expected}&lt;/P&gt;&lt;/DIV&gt;</parameterizedString>
                                    <description/>
                                </step>'''
                            steps_xml += '</steps>'
                            
                            update_data = [
                                {"op": "add", "path": "/fields/Microsoft.VSTS.TCM.Steps", "value": steps_xml}
                            ]
                            update_url = f"{base_url}/wit/workitems/{work_item_id}?api-version=7.0"
                            requests.patch(update_url, headers=headers, json=update_data, timeout=30)
                        
                        all_results.append({
                            'Batch': batch_num + 1 if num_batches > 1 else 'N/A',
                            'Title': tc['title'],
                            'Status': 'Success',
                            'Work Item ID': work_item_id,
                            'Steps': len(tc['steps']),
                            'Timestamp': datetime.now().strftime('%H:%M:%S')
                        })
                    else:
                        error_msg = response.text[:100] if response.text else 'Unknown error'
                        all_results.append({
                            'Batch': batch_num + 1 if num_batches > 1 else 'N/A',
                            'Title': tc['title'],
                            'Status': f'Failed ({response.status_code})',
                            'Work Item ID': None,
                            'Steps': len(tc['steps']),
                            'Timestamp': datetime.now().strftime('%H:%M:%S')
                        })
                        print(f"Upload failed for {tc['title']}: {error_msg}")
                        
                except Exception as e:
                    all_results.append({
                        'Batch': batch_num + 1 if num_batches > 1 else 'N/A',
                        'Title': tc['title'],
                        'Status': f'Error: {str(e)[:50]}',
                        'Work Item ID': None,
                        'Steps': len(tc['steps']),
                        'Timestamp': datetime.now().strftime('%H:%M:%S')
                    })
                    print(f"Exception for {tc['title']}: {str(e)}")
                
                # Rate limiting - be nice to the API
                time.sleep(0.5)
            
            # Pause between batches (if multiple batches)
            if num_batches > 1 and batch_num < num_batches - 1:
                upload_progress_info.set({
                    'status': 'running',
                    'current': end_idx,
                    'total': total,
                    'batch': batch_num + 1,
                    'total_batches': num_batches,
                    'message': f'Completed batch {batch_num + 1}/{num_batches}. Pausing before next batch...'
                })
                time.sleep(2)  # Brief pause between batches
        
        upload_results_data.set(pd.DataFrame(all_results))
        
        completion_msg = f'Upload complete! Processed {total} test cases'
        if num_batches > 1:
            completion_msg += f' across {num_batches} batches'
        
        upload_progress_info.set({
            'status': 'complete',
            'current': total,
            'total': total,
            'batch': num_batches if num_batches > 1 else None,
            'total_batches': num_batches if num_batches > 1 else None,
            'message': completion_msg
        })
    
    # ========================================================================
    # UPLOAD PROGRESS & RESULTS
    # ========================================================================
    
    @output
    @render.ui
    def upload_progress():
        """Display upload progress"""
        progress_info = upload_progress_info.get()
        results_df = upload_results_data.get()
        
        if progress_info is None and results_df is None:
            return ui.div(
                ui.tags.p("Ready to upload. Click 'Upload to Azure DevOps' to start.", 
                         class_="text-info")
            )
        
        elements = []
        
        if progress_info:
            if progress_info['status'] == 'info':
                elements.append(ui.tags.div(
                    ui.tags.h5("ℹ️ Batch Processing", class_="text-info"),
                    ui.tags.p(progress_info['message']),
                    class_="info-box"
                ))
            elif progress_info['status'] == 'running':
                progress_pct = (progress_info['current'] / progress_info['total']) * 100
                
                # Show batch info if applicable
                if progress_info.get('batch') and progress_info.get('total_batches'):
                    elements.append(ui.tags.h5(
                        f"Batch {progress_info['batch']}/{progress_info['total_batches']} - " +
                        f"Progress: {progress_info['current']}/{progress_info['total']}"
                    ))
                else:
                    elements.append(ui.tags.h5(f"Progress: {progress_info['current']}/{progress_info['total']}"))
                
                elements.append(ui.tags.div(
                    ui.tags.div(
                        style=f"width: {progress_pct}%; height: 25px; background-color: #007bff; text-align: center; color: white; line-height: 25px;",
                        children=f"{progress_pct:.1f}%"
                    ),
                    style="width: 100%; background-color: #e9ecef; border-radius: 5px;"
                ))
                elements.append(ui.tags.p(f"Current: {progress_info['message']}"))
            elif progress_info['status'] == 'complete':
                elements.append(ui.tags.h4("✅ Upload Complete!", class_="text-success"))
                if progress_info.get('message'):
                    elements.append(ui.tags.p(progress_info['message']))
            elif progress_info['status'] == 'error':
                elements.append(ui.tags.h4("❌ Error", class_="text-danger"))
                elements.append(ui.tags.p(progress_info['message']))
        
        if results_df is not None:
            successful = len(results_df[results_df['Status'] == 'Success'])
            failed = len(results_df) - successful
            
            elements.append(ui.tags.hr())
            elements.append(ui.tags.h5("Summary:"))
            elements.append(ui.tags.p(f"Total: {len(results_df)}"))
            elements.append(ui.tags.p(
                ui.tags.span(f"✅ Successful: {successful}", class_="success-badge" if successful > 0 else ""),
                " ",
                ui.tags.span(f"❌ Failed: {failed}", class_="error-badge" if failed > 0 else "")
            ))
            
            # Show batch breakdown if multiple batches were used
            if 'Batch' in results_df.columns and results_df['Batch'].nunique() > 1:
                elements.append(ui.tags.hr())
                elements.append(ui.tags.h5("Batch Breakdown:"))
                batch_summary = results_df.groupby('Batch').agg({
                    'Status': lambda x: (x == 'Success').sum()
                }).reset_index()
                batch_summary.columns = ['Batch', 'Successful']
                batch_summary['Total'] = results_df.groupby('Batch').size().values
                
                for _, row in batch_summary.iterrows():
                    elements.append(ui.tags.p(
                        f"Batch {row['Batch']}: {row['Successful']}/{row['Total']} successful"
                    ))
        
        return ui.div(*elements)
    
    @output
    @render.table
    def upload_results():
        """Display detailed upload results"""
        results_df = upload_results_data.get()
        if results_df is None:
            return pd.DataFrame({"Message": ["No upload results yet"]})
        return results_df


# ============================================================================
# CREATE AND RUN APP
# ============================================================================

app = App(app_ui, server)