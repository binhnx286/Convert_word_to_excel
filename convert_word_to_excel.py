import tkinter as tk
from tkinter import filedialog, messagebox
import docx
import pandas as pd
import re

def read_questions_from_docx(docx_file):
    doc = docx.Document(docx_file)
    questions = []
    current_question = None
    current_answers = {"A": "", "B": "", "C": "", "D": ""}
    current_correct_answer = ""  # Biến lưu trữ câu trả lời đúng
    
    for para in doc.paragraphs:
        text = para.text.strip()
        if text.startswith("Câu "):
            # Finish processing previous question
            if current_question is not None:
                question_data = [current_question] + list(current_answers.values()) + [current_correct_answer]
                questions.append(question_data)
            
            # Start new question
            current_question = re.sub(r'^Câu \d+ ?: ?', '', text)  # Remove "Câu x:" or "Câu x:"
            current_answers = {"A": "", "B": "", "C": "", "D": ""}
            current_correct_answer = ""  # Reset câu trả lời đúng
            
        elif text.startswith("A."):
            current_answers["A"] = text[3:].strip()
        elif text.startswith("B."):
            current_answers["B"] = text[3:].strip()
        elif text.startswith("C."):
            current_answers["C"] = text[3:].strip()
        elif text.startswith("D."):
            current_answers["D"] = text[3:].strip()
        elif text.startswith("Cr:"):
            current_correct_answer = text.split(":", 1)[1].strip()
    
    # Append last question and answers
    if current_question is not None:
        question_data = [current_question] + list(current_answers.values()) + [current_correct_answer]
        questions.append(question_data)
    
    return questions

def write_questions_to_excel(questions, excel_file):
    df = pd.DataFrame(questions, columns=['Questions', 'Answer A', 'Answer B', 'Answer C', 'Answer D', 'Correct Answer'])
    df.to_excel(excel_file, index=False)
    messagebox.showinfo("Success", f"Questions extracted and saved to '{excel_file}' successfully.")

def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx"), ("All files", "*.*")])
    if file_path:
        questions = read_questions_from_docx(file_path)
        excel_file = file_path.rsplit('.', 1)[0] + '.xlsx'
        write_questions_to_excel(questions, excel_file)

# Create the main window
root = tk.Tk()
root.title("Convert Word to Excel")

# Create a button to select file
select_button = tk.Button(root, text="Select Word File", command=select_file)
select_button.pack(pady=20)

# Run the main loop
root.mainloop()
