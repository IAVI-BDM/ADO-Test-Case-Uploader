"""
Azure DevOps Test Case Uploader
A Python Shiny application for processing and uploading test cases to Azure DevOps

Author: Biostatistics & Data Science
Version: 2.0
"""

from shiny import App, ui, render, reactive
import pandas as pd
import requests
import json
import base64
import time
from datetime import datetime

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
            ui.input_text("iteration_path", "Iteration Path", 
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
            
            ui.nav_panel("📝 Test Cases",
                ui.card(
                    ui.card_header("Processed Test Cases Summary"),
                    ui.output_ui("test_cases_summary")
                ),
                ui.card(
                    ui.card_header("Test Case Examples"),
                    ui.output_ui("test_cases_display")
                )
            ),
            
            ui.nav_panel("🔄 Upload Progress",
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
                    
                    **Optional CSV Columns:**
                    - `Custom.FieldName` (for Field Level)
                    - `Custom.EditCheckName` (for Edit Check Level)
                    - `Area Path`
                    - `Iteration Path`
                    - `State`
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
            
            # Validate required columns
            required_cols = ['Custom.TestCaseClassification', 'Custom.FormName']
            missing_cols = [col for col in required_cols if col not in df.columns]
            
            if missing_cols:
                print(f"Warning: Missing required columns: {missing_cols}")
            
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
        
        return ui.div(*stats)
    
    @output
    @render.table
    def data_preview():
        """Show preview of uploaded data"""
        df = uploaded_data.get()
        if df is None:
            return pd.DataFrame({"Message": ["No data loaded"]})
        return df.head(10)
    
    # ========================================================================
    # TEST CASE PROCESSING
    # ========================================================================
    
    @reactive.Effect
    @reactive.event(input.process_btn)
    def process_test_cases():
        """Process CSV into test case structure"""
        df = uploaded_data.get()
        if df is None:
            return
        
        test_cases = []
        
        # Group by Form Name
        unique_forms = df['Custom.FormName'].dropna().unique()
        
        for form_name in unique_forms:
            form_data = df[df['Custom.FormName'] == form_name]
            
            # Process Form Level items (standalone test cases)
            form_level = form_data[form_data['Custom.TestCaseClassification'] == 'Form Level']
            for _, row in form_level.iterrows():
                test_cases.append({
                    'type': 'standalone',
                    'title': f"Form - {form_name}",
                    'form_name': form_name,
                    'classification': 'Form Level',
                    'description': str(row.get('Custom.FieldorEditCheckText', '')),
                    'area_path': str(row.get('Area Path', '')),
                    'iteration_path': str(row.get('Iteration Path', '')),
                    'state': str(row.get('State', 'Design')),
                    'steps': []
                })
            
            # Process Field Level items (group into Field Reviews test case)
            field_level = form_data[form_data['Custom.TestCaseClassification'] == 'Field Level']
            if len(field_level) > 0:
                steps = []
                for idx, row in field_level.iterrows():
                    field_name = str(row.get('Custom.FieldName', ''))
                    field_text = str(row.get('Custom.FieldorEditCheckText', ''))
                    steps.append({
                        'step_number': len(steps) + 1,
                        'action': f"Test field: {field_name}",
                        'expected': f"Field '{field_name}' validates correctly. {field_text}",
                        'field_name': field_name
                    })
                
                first_row = field_level.iloc[0]
                test_cases.append({
                    'type': 'field_reviews',
                    'title': f"{form_name} - Field Reviews",
                    'form_name': form_name,
                    'classification': 'Field Level',
                    'description': f"Field-level validation for form {form_name}. Total fields: {len(steps)}",
                    'area_path': str(first_row.get('Area Path', '')),
                    'iteration_path': str(first_row.get('Iteration Path', '')),
                    'state': str(first_row.get('State', 'Design')),
                    'steps': steps
                })
            
            # Process Edit Check Level items (group into Edit Check Reviews test case)
            edit_check_level = form_data[form_data['Custom.TestCaseClassification'] == 'Edit Check Level']
            if len(edit_check_level) > 0:
                steps = []
                for idx, row in edit_check_level.iterrows():
                    edit_check_name = str(row.get('Custom.EditCheckName', ''))
                    edit_check_text = str(row.get('Custom.FieldorEditCheckText', ''))
                    steps.append({
                        'step_number': len(steps) + 1,
                        'action': f"Test edit check: {edit_check_name}",
                        'expected': f"Edit check '{edit_check_name}' functions correctly. {edit_check_text}",
                        'edit_check_name': edit_check_name
                    })
                
                first_row = edit_check_level.iloc[0]
                test_cases.append({
                    'type': 'edit_check_reviews',
                    'title': f"{form_name} - Edit Check Reviews",
                    'form_name': form_name,
                    'classification': 'Edit Check Level',
                    'description': f"Edit check validation for form {form_name}. Total checks: {len(steps)}",
                    'area_path': str(first_row.get('Area Path', '')),
                    'iteration_path': str(first_row.get('Iteration Path', '')),
                    'state': str(first_row.get('State', 'Design')),
                    'steps': steps
                })
        
        processed_test_cases.set(test_cases)
    
    # ========================================================================
    # TEST CASE DISPLAY
    # ========================================================================
    
    @output
    @render.ui
    def test_cases_summary():
        """Show summary of processed test cases"""
        test_cases = processed_test_cases.get()
        if test_cases is None:
            return ui.div(
                ui.tags.p("⚠️ No test cases processed. Click 'Process CSV' first.", 
                         class_="text-warning")
            )
        
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
            'field_reviews': '📝 Field Reviews',
            'edit_check_reviews': '✅ Edit Check Reviews'
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
        summary.append(ui.tags.div(
            ui.tags.p(
                "✅ Processed file ready! ",
                ui.tags.br(),
                "Click 'Download Processed CSV' to export this structure."
            ),
            class_="info-box"
        ))
        
        return ui.div(*summary)
    
    @output
    @render.ui
    def test_cases_display():
        """Display sample test cases"""
        test_cases = processed_test_cases.get()
        if test_cases is None:
            return ui.p("No test cases to display.")
        
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
            
            content = ui.div(
                ui.tags.p(ui.tags.strong("Title: "), tc['title']),
                ui.tags.p(ui.tags.strong("Type: "), tc['type'].replace('_', ' ').title()),
                ui.tags.p(ui.tags.strong("Form: "), tc['form_name']),
                ui.tags.p(ui.tags.strong("Total Steps: "), str(len(tc['steps']))),
                ui.tags.hr(),
                step_info if tc['steps'] else ui.tags.em("No steps (standalone test case)")
            )
            
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
        """Download processed test cases as CSV"""
        test_cases = processed_test_cases.get()
        
        if test_cases is None:
            # Return empty CSV with message
            return pd.DataFrame({"Message": ["No processed test cases available. Please process CSV first."]}).to_csv(index=False)
        
        # Convert test cases to flat CSV format
        rows = []
        
        for tc in test_cases:
            # If test case has no steps (standalone)
            if not tc['steps']:
                rows.append({
                    'Test Case Title': tc['title'],
                    'Test Case Type': tc['type'],
                    'Form Name': tc['form_name'],
                    'Classification': tc['classification'],
                    'Description': tc['description'],
                    'Area Path': tc['area_path'],
                    'Iteration Path': tc['iteration_path'],
                    'State': tc['state'],
                    'Total Steps': 0,
                    'Step Number': None,
                    'Step Action': None,
                    'Step Expected': None
                })
            else:
                # Create one row per step
                for step in tc['steps']:
                    rows.append({
                        'Test Case Title': tc['title'],
                        'Test Case Type': tc['type'],
                        'Form Name': tc['form_name'],
                        'Classification': tc['classification'],
                        'Description': tc['description'],
                        'Area Path': tc['area_path'],
                        'Iteration Path': tc['iteration_path'],
                        'State': tc['state'],
                        'Total Steps': len(tc['steps']),
                        'Step Number': step['step_number'],
                        'Step Action': step['action'],
                        'Step Expected': step['expected']
                    })
        
        df = pd.DataFrame(rows)
        return df.to_csv(index=False)
    
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
        test_cases = processed_test_cases.get()
        organization = input.organization()
        project = input.project()
        pat_token = input.pat_token()
        dry_run = input.dry_run()
        
        if not all([test_cases, organization, project, pat_token]):
            upload_progress_info.set({
                'status': 'error',
                'message': 'Missing required configuration. Please fill all fields.'
            })
            return
        
        base_url = f"https://dev.azure.com/{organization}/{project}/_apis"
        headers = create_auth_header(pat_token)
        
        results = []
        total = len(test_cases)
        
        for idx, tc in enumerate(test_cases, 1):
            # Update progress
            upload_progress_info.set({
                'status': 'running',
                'current': idx,
                'total': total,
                'message': f'Processing: {tc["title"]}'
            })
            
            try:
                if dry_run:
                    # Simulate upload for dry run
                    time.sleep(0.1)
                    results.append({
                        'Title': tc['title'],
                        'Status': 'Dry Run',
                        'Work Item ID': 'N/A',
                        'Steps': len(tc['steps']),
                        'Timestamp': datetime.now().strftime('%H:%M:%S')
                    })
                    continue
                
                # Build work item data
                area_path = input.area_path() or tc['area_path']
                iteration_path = input.iteration_path() or tc['iteration_path']
                
                work_item_data = [
                    {"op": "add", "path": "/fields/System.Title", "value": tc['title']},
                    {"op": "add", "path": "/fields/System.State", "value": tc['state']},
                    {"op": "add", "path": "/fields/System.Description", "value": tc['description']}
                ]
                
                if area_path:
                    work_item_data.append({"op": "add", "path": "/fields/System.AreaPath", "value": area_path})
                if iteration_path:
                    work_item_data.append({"op": "add", "path": "/fields/System.IterationPath", "value": iteration_path})
                
                # Add custom fields (if they exist in your project template)
                try:
                    work_item_data.append({"op": "add", "path": "/fields/Custom.TestCaseClassification", "value": tc['classification']})
                    work_item_data.append({"op": "add", "path": "/fields/Custom.FormName", "value": tc['form_name']})
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
                    
                    results.append({
                        'Title': tc['title'],
                        'Status': 'Success',
                        'Work Item ID': work_item_id,
                        'Steps': len(tc['steps']),
                        'Timestamp': datetime.now().strftime('%H:%M:%S')
                    })
                else:
                    error_msg = response.text[:100] if response.text else 'Unknown error'
                    results.append({
                        'Title': tc['title'],
                        'Status': f'Failed ({response.status_code})',
                        'Work Item ID': None,
                        'Steps': len(tc['steps']),
                        'Timestamp': datetime.now().strftime('%H:%M:%S')
                    })
                    print(f"Upload failed for {tc['title']}: {error_msg}")
                    
            except Exception as e:
                results.append({
                    'Title': tc['title'],
                    'Status': f'Error: {str(e)[:50]}',
                    'Work Item ID': None,
                    'Steps': len(tc['steps']),
                    'Timestamp': datetime.now().strftime('%H:%M:%S')
                })
                print(f"Exception for {tc['title']}: {str(e)}")
            
            # Rate limiting - be nice to the API
            time.sleep(0.5)
        
        upload_results_data.set(pd.DataFrame(results))
        upload_progress_info.set({
            'status': 'complete',
            'current': total,
            'total': total,
            'message': 'Upload complete!'
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
            if progress_info['status'] == 'running':
                progress_pct = (progress_info['current'] / progress_info['total']) * 100
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