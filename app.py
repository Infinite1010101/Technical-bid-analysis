from flask import Flask, render_template, request, jsonify, send_file
import pandas as pd
import os
import io
from PyPDF2 import PdfReader # PDF files padhne ke liye
from docx import Document # DOCX files padhne ke liye

# Agar aap OpenAI API use kar rahe hain, toh yeh line uncomment karein
# import openai
# import json

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads' # Temporary files store karne ke liye folder
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True) # Folder banayein agar nahi hai

# --- Home Page Route ---
@app.route('/')
def index():
    # Yeh function 'templates' folder se index.html file ko load karega
    return render_template('index.html')

# --- File Upload aur Comparison Route ---
@app.route('/upload_and_compare', methods=['POST'])
def upload_and_compare():
    # Check karein ki files upload hui hain ya nahi
    if 'tech_spec_file' not in request.files or 'vendor_offers_file' not in request.files:
        return jsonify({"error": "Technical Specification file or Vendor Offers file is missing."}), 400

    tech_spec_file = request.files['tech_spec_file']
    vendor_offers_file = request.files['vendor_offers_file']
    project_name = request.form.get('project_name', 'Untitled Project') # Project name form se lein

    if tech_spec_file.filename == '' or vendor_offers_file.filename == '':
        return jsonify({"error": "No selected file for Technical Specification or Vendor Offers."}), 400

    if tech_spec_file and vendor_offers_file:
        # Files ko temporary folder mein save karein
        tech_spec_path = os.path.join(app.config['UPLOAD_FOLDER'], tech_spec_file.filename)
        vendor_offers_path = os.path.join(app.config['UPLOAD_FOLDER'], vendor_offers_file.filename)
        tech_spec_file.save(tech_spec_path)
        vendor_offers_file.save(vendor_offers_path)

        # --- Core Logic: Files Padhein, Process Karein, Compare Karein ---
        try:
            # 1. Technical Specifications Padhna
            tech_spec_content = read_file_content(tech_spec_path)

            # 2. Vendor Offers Padhna
            vendor_data = read_vendor_offers_excel(vendor_offers_path)

            # 3. AI/NLP Processing (Yeh sabse mushkil hissa hai)
            # Yahan aap AI/NLP API ko call karenge
            comparison_results = perform_nlp_comparison(tech_spec_content, vendor_data)

            # 4. Excel Output Banana
            output_excel_buffer = generate_excel_report(project_name, comparison_results)

            # Temporary files delete karein (optional, acchi practice hai)
            os.remove(tech_spec_path)
            os.remove(vendor_offers_path)

            # Excel file ko user ko download karne ke liye bhejein
            return send_file(
                output_excel_buffer,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                as_attachment=True,
                download_name=f'{project_name}_Technical_Comparison.xlsx'
            )

        except Exception as e:
            # Agar koi error aata hai toh user ko batayein
            return jsonify({"error": f"Processing failed: {str(e)}. Please check file formats and content."}), 500

# --- Helper Functions (Yeh functions files ko padhne aur Excel banane mein madad karte hain) ---

def read_file_content(filepath):
    """
    File content ko padhta hai (TXT, XLSX, PDF, DOCX).
    """
    if filepath.endswith('.txt'):
        with open(filepath, 'r', encoding='utf-8') as f:
            return f.read()
    elif filepath.endswith('.xlsx') or filepath.endswith('.xls'):
        # Agar tech spec Excel mein hai, toh pehli sheet ke saare text ko combine karein
        df = pd.read_excel(filepath)
        # Aapko yahan decide karna hoga ki Excel mein tech spec kis column mein hai
        # Abhi ke liye, saare columns ke text ko join kar rahe hain
        return df.to_string(index=False)
    elif filepath.endswith('.pdf'):
        reader = PdfReader(filepath)
        text = ""
        for page in reader.pages:
            text += page.extract_text() + "\n"
        return text
    elif filepath.endswith('.docx'):
        doc = Document(filepath)
        text = ""
        for para in doc.paragraphs:
            text += para.text + "\n"
        return text
    else:
        raise ValueError("Unsupported file type for technical specification. Please use .txt, .pdf, .docx, .xls, or .xlsx.")

def read_vendor_offers_excel(filepath):
    """
    Vendor offers ko Excel file se padhta hai.
    Assumes ki Excel mein 'Vendor Name' aur 'Offer Details' (ya similar) columns hain.
    """
    df = pd.read_excel(filepath)
    vendor_offers = []
    for index, row in df.iterrows():
        vendor_offers.append({
            "name": row.get('Vendor Name', f'Vendor {index+1}'), # Agar 'Vendor Name' column nahi hai
            "offer_text": row.get('Offer Details', '') # Agar 'Offer Details' column nahi hai
            # Agar aapke Excel mein aur bhi columns hain jo offer ka part hain, toh unhe yahan combine karein
            # Example: "offer_text": f"{row.get('Offer Details', '')} {row.get('Technical Compliance', '')}"
        })
    if not vendor_offers:
        raise ValueError("Vendor offers Excel file is empty or does not contain expected columns like 'Vendor Name' or 'Offer Details'.")
    return vendor_offers

def perform_nlp_comparison(tech_spec_content, vendor_data):
    """
    Yeh function AI/NLP ka core hai.
    Yahan aap OpenAI ya kisi aur NLP API ka use karenge.
    """
    # --- IMPORTANT: Yahan aapko apni AI/NLP API key set karni hogi ---
    # Agar aap OpenAI use kar rahe hain:
    # import openai
    # import json
    # openai.api_key = os.getenv("OPENAI_API_KEY") # Best practice: environment variable se lein
    # Agar aap seedha code mein likh rahe hain (sirf testing ke liye):
    # openai.api_key = "YOUR_OPENAI_API_KEY_HERE" # Apni actual API key yahan daalein

    # Agar API key set nahi hai toh error dein
    # if not openai.api_key:
    #     raise ValueError("OpenAI API key not set. Please set it as an environment variable or directly in the code.")

    # --- Step 1: Technical Specifications se Requirements Extract Karna ---
    # Yeh hissa OpenAI API call karega
    extracted_requirements = []
    # try:
    #     requirements_prompt = f"""
    #     Extract all key technical requirements and their specific details/values from the following text.
    #     Focus on quantifiable or clearly defined criteria.
    #     Present the output as a JSON array of objects, where each object has 'requirement' and 'details' keys.
    #     Example:
    #     [
    #       {{"requirement": "Processor", "details": "Intel Core i7-12700H or equivalent"}},
    #       {{"requirement": "RAM", "details": "32GB DDR5"}}
    #     ]
    #     Technical Specification:
    #     {tech_spec_content}
    #     """
    #     response = openai.chat.completions.create(
    #         model="gpt-4o", # Ya "gpt-3.5-turbo"
    #         messages=[
    #             {"role": "system", "content": "You are an expert technical analyst. Extract requirements accurately."},
    #             {"role": "user", "content": requirements_prompt}
    #         ],
    #         response_format={"type": "json_object"}
    #     )
    #     extracted_requirements_str = response.choices[0].message.content
    #     # OpenAI kabhi-kabhi JSON ko ek dictionary ke andar deta hai, jaise {'requirements': [...]}
    #     parsed_json = json.loads(extracted_requirements_str)
    #     extracted_requirements = parsed_json.get('requirements', parsed_json) # 'requirements' key se lein ya seedha JSON ko
    # except Exception as e:
    #     print(f"Error extracting requirements with OpenAI: {e}")
    #     # Fallback: Agar API fail ho toh simple line-by-line requirements lein
    #     extracted_requirements = [{"requirement": req.strip(), "details": ""} for req in tech_spec_content.split('\n') if req.strip()]

    # --- Fallback/Simple Logic (Agar aap AI API use nahi kar rahe ya test kar rahe hain) ---
    # Agar aap AI API use nahi kar rahe, toh yahan aap simple keyword matching ya rules bana sakte hain.
    # Abhi ke liye, hum technical specification ki har line ko ek requirement maan rahe hain.
    extracted_requirements = [{"requirement": req.strip(), "details": ""} for req in tech_spec_content.split('\n') if req.strip()]
    if not extracted_requirements:
        raise ValueError("Could not extract any requirements from the technical specification. Please ensure it contains clear text.")


    # --- Step 2: Har Vendor Offer ko Compare Karna ---
    final_comparison_data = []
    for vendor in vendor_data:
        vendor_name = vendor['name']
        vendor_offer_text = vendor['offer_text']

        # Yeh hissa bhi OpenAI API call karega
        # try:
        #     vendor_comparison_prompt = f"""
        #     You are comparing a vendor's technical offer against a list of specific requirements.
        #     For each requirement, determine if the vendor's offer meets, partially meets, or does not meet it.
        #     Provide a brief explanation and relevant text from the vendor's offer as evidence.
        #     Also, list any additional significant features the vendor offers that are not in the requirements.
        #     Technical Requirements:
        #     {json.dumps(extracted_requirements, indent=2)}
        #     Vendor Offer from {vendor_name}:
        #     {vendor_offer_text}
        #     Present the output as a JSON object with the following structure:
        #     {{
        #       "vendor_name": "Vendor A",
        #       "overall_status": "Met All Requirements" or "Partially Met" or "Did Not Meet",
        #       "detailed_comparison": [
        #         {{
        #           "requirement": "Processor",
        #           "status": "Met",
        #           "explanation": "Vendor offers Intel i9-13900H, which exceeds the i7 requirement.",
        #           "evidence": "Processor: Intel Core i9-13900H"
        #         }},
        #         ...
        #       ],
        #       "additional_features": ["Extended warranty", "On-site support"]
        #     }}
        #     """
        #     response = openai.chat.completions.create(
        #         model="gpt-4o", # Ya "gpt-3.5-turbo"
        #         messages=[
        #             {"role": "system", "content": "You are a meticulous technical bid evaluator."},
        #             {"role": "user", "content": vendor_comparison_prompt}
        #         ],
        #         response_format={"type": "json_object"}
        #     )
        #     vendor_analysis_str = response.choices[0].message.content
        #     vendor_analysis = json.loads(vendor_analysis_str)
        #     final_comparison_data.append(vendor_analysis)
        # except Exception as e:
        #     print(f"Error analyzing vendor {vendor_name} with OpenAI: {e}")
        #     final_comparison_data.append({
        #         "vendor_name": vendor_name,
        #         "overall_status": "Error during AI analysis",
        #         "detailed_comparison": [],
        #         "additional_features": []
        #     })

        # --- Fallback/Simple Logic (Agar aap AI API use nahi kar rahe ya test kar rahe hain) ---
        # Yahan hum ek bahut hi simple keyword matching kar rahe hain.
        # Real AI/NLP yahan bahut complex analysis karega.
        vendor_result = {
            "vendor_name": vendor_name,
            "overall_status": "Pending Analysis", # Default status
            "detailed_comparison": [],
            "additional_features": []
        }
        met_count = 0
        for req_obj in extracted_requirements:
            req_text = req_obj['requirement'].lower()
            status = "Not Met"
            explanation = "Not found in offer."
            evidence = ""

            if req_text in vendor_offer_text.lower():
                status = "Met"
                explanation = "Found in offer."
                met_count += 1
                # Yahan aap offer text se relevant snippet nikal sakte hain
                start_idx = vendor_offer_text.lower().find(req_text)
                if start_idx != -1:
                    evidence = vendor_offer_text[start_idx:start_idx + len(req_text) + 20] + "..." # Small snippet

            vendor_result["detailed_comparison"].append({
                "requirement": req_obj['requirement'],
                "status": status,
                "explanation": explanation,
                "evidence": evidence
            })

        if met_count == len(extracted_requirements) and len(extracted_requirements) > 0:
            vendor_result["overall_status"] = "Met All Requirements"
        elif met_count > 0:
            vendor_result["overall_status"] = "Partially Met"
        elif len(extracted_requirements) > 0:
            vendor_result["overall_status"] = "Did Not Meet"
        else:
            vendor_result["overall_status"] = "No Requirements Defined"

        final_comparison_data.append(vendor_result)

    return final_comparison_data


def generate_excel_report(project_name, comparison_results):
    """
    Comparison results se Excel report banata hai.
    """
    # DataFrame banane ke liye data ko flat karein
    rows_for_excel = []
    # Pehle saare possible requirements collect karein
    all_requirements = set()
    for res in comparison_results:
        for detail in res.get("detailed_comparison", []):
            all_requirements.add(detail["requirement"])
    sorted_requirements = sorted(list(all_requirements)) # Requirements ko alphabetical order mein sort karein

    for res in comparison_results:
        row = {
            "Project Name": project_name,
            "Vendor": res.get("vendor_name", "N/A"),
            "Overall Status": res.get("overall_status", "N/A"),
            "Additional Features": ", ".join(res.get("additional_features", []))
        }
        # Har requirement ke liye columns add karein
        req_details_map = {d["requirement"]: d for d in res.get("detailed_comparison", [])}
        for req_name in sorted_requirements:
            detail = req_details_map.get(req_name, {})
            row[f"{req_name} - Status"] = detail.get("status", "N/A")
            row[f"{req_name} - Explanation"] = detail.get("explanation", "N/A")
            row[f"{req_name} - Evidence"] = detail.get("evidence", "N/A")
        rows_for_excel.append(row)

    df = pd.DataFrame(rows_for_excel)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Technical Comparison', index=False)

        # Optional: Excel sheet ko thoda format karein
        worksheet = writer.sheets['Technical Comparison']
        for column in worksheet.columns:
            max_length = 0
            column_name = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2)
            worksheet.column_dimensions[column_name].width = adjusted_width

    output.seek(0) # Buffer ke shuruat mein wapas jayein
    return output

# --- Application Run Karna ---
if __name__ == '__main__':
    # debug=True development ke liye accha hai, production mein False karein
    app.run(debug=True)