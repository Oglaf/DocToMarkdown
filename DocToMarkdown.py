import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox
import os
import sys
import threading
import configparser
from pathlib import Path
import subprocess
import re
import shutil

# Attempt to import dependencies and handle the error if they are not installed.
try:
    from cryptography.fernet import Fernet
    from openai import AzureOpenAI
except ImportError:
    # This will be handled in the __main__ block before the GUI starts.
    pass

# --- Configuration ---
CONFIG_FILE = 'config.ini'
KEY_FILE = 'key.key'

def generate_key():
    """Generates a new encryption key and saves it to a file."""
    key = Fernet.generate_key()
    with open(KEY_FILE, "wb") as key_file:
        key_file.write(key)
    return key

def load_key():
    """Loads the encryption key from the key file, or generates a new one."""
    if not os.path.exists(KEY_FILE):
        return generate_key()
    with open(KEY_FILE, "rb") as key_file:
        return key_file.read()

def encrypt(data, fernet):
    """Encrypts data using the provided Fernet instance."""
    if not data: return ''
    return fernet.encrypt(data.encode()).decode()

def decrypt(token, fernet):
    """Decrypts a token using the provided Fernet instance."""
    if not token: return ''
    return fernet.decrypt(token.encode()).decode()

def save_config(vars_dict):
    """Saves the current configuration to the INI file, encrypting sensitive fields."""
    config = configparser.ConfigParser()
    key = load_key()
    fernet = Fernet(key)

    config['Paths'] = {
        'pandoc_executable_path': vars_dict['pandoc_path'].get(),
        'last_wiki_root': vars_dict['wiki_root_path'].get()
    }
    config['AzureOpenAI'] = {
        'endpoint': vars_dict['gpt_endpoint'].get(),
        'key': encrypt(vars_dict['gpt_key'].get(), fernet),
        'deployment': vars_dict['gpt_deployment'].get()
    }
    config['Settings'] = {
        'post_process_enabled': str(vars_dict['post_process_var'].get()),
        'ai_prompt': vars_dict['prompt_text'].get("1.0", "end-1c")
    }

    with open(CONFIG_FILE, 'w') as configfile:
        config.write(configfile)
    messagebox.showinfo("Success", "Configuration saved successfully!")

def load_config(vars_dict):
    """Loads configuration from the INI file and populates the UI fields."""
    if not os.path.exists(CONFIG_FILE):
        return
    config = configparser.ConfigParser()
    config.read(CONFIG_FILE)
    key = load_key()
    fernet = Fernet(key)

    vars_dict['pandoc_path'].set(config.get('Paths', 'pandoc_executable_path', fallback='pandoc'))
    vars_dict['wiki_root_path'].set(config.get('Paths', 'last_wiki_root', fallback=''))
    
    try:
        vars_dict['gpt_endpoint'].set(config.get('AzureOpenAI', 'endpoint', fallback=''))
        vars_dict['gpt_key'].set(decrypt(config.get('AzureOpenAI', 'key', fallback=''), fernet))
        vars_dict['gpt_deployment'].set(config.get('AzureOpenAI', 'deployment', fallback=''))

        vars_dict['post_process_var'].set(config.getboolean('Settings', 'post_process_enabled', fallback=False))
        vars_dict['prompt_text'].delete("1.0", tk.END)
        vars_dict['prompt_text'].insert("1.0", config.get('Settings', 'ai_prompt', fallback=''))
    except Exception:
        pass

def post_process_with_ai(file_path, gpt_endpoint, gpt_key, gpt_deployment, prompt, output_widget):
    """Sends the markdown content to Azure OpenAI for post-processing."""
    output_widget.insert(tk.END, "\nStarting AI post-processing...\n")
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            markdown_content = f.read()

        client = AzureOpenAI(api_version="2024-02-01", azure_endpoint=gpt_endpoint, api_key=gpt_key)
        
        response = client.chat.completions.create(
            model=gpt_deployment,
            messages=[
                {"role": "system", "content": "You are an expert markdown editor..."},
                {"role": "user", "content": f"PROMPT:\n---\n{prompt}\n---\n\nMARKDOWN:\n---\n{markdown_content}"}
            ]
        )
        
        processed_content = response.choices[0].message.content
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(processed_content)
        output_widget.insert(tk.END, "AI post-processing completed successfully.\n")
    except Exception as e:
        output_widget.insert(tk.END, f"--- ERROR during AI post-processing: {e} ---\n")

def run_conversion_logic(pandoc_exe, file_path, output_dir, attachments_dir_name, wiki_root_dir, post_process, ai_params, output_widget):
    """Runs the full conversion and optional post-processing pipeline."""
    try:
        output_widget.insert(tk.END, "Preparing for Pandoc conversion...\n")
        
        output_path = Path(output_dir)
        attachments_path = Path(wiki_root_dir) / attachments_dir_name
        output_path.mkdir(parents=True, exist_ok=True)
        attachments_path.mkdir(parents=True, exist_ok=True)
        
        input_filename_stem = Path(file_path).stem
        sanitized_stem = input_filename_stem.replace(' ', '-')
        output_file_name = f"{sanitized_stem}.md"
        output_file_path = output_path / output_file_name

        command = [
            pandoc_exe, str(file_path),
            "-t", "markdown",
            "--extract-media", str(attachments_path),
            "-o", output_file_name
        ]

        output_widget.insert(tk.END, f"Working directory: {output_path}\nRunning command: {' '.join(command)}\n\n")

        startupinfo = None
        if sys.platform == "win32":
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW

        process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, startupinfo=startupinfo, encoding='utf-8', cwd=output_path)
        process.wait()

        if process.returncode == 0:
            output_widget.insert(tk.END, "\nPandoc conversion successful. Cleaning up Markdown...\n")
            
            media_subfolder = attachments_path / "media"
            if media_subfolder.is_dir():
                for item in os.listdir(media_subfolder):
                    shutil.move(str(media_subfolder / item), str(attachments_path / item))
                media_subfolder.rmdir()

            with open(output_file_path, 'r+', encoding='utf-8') as f:
                content = f.read()
                content = re.sub(r'(\!\[.*?\]\()' + re.escape(str(attachments_path).replace('\\', '/')) + r'/(.*?)\)', r'\1../' + attachments_dir_name + r'/\2)', content)
                content = re.sub(r'(\!\[.*?\]\()' + attachments_dir_name + r'/(.*?)\)', r'![](../' + attachments_dir_name + r'/\2)', content)
                content = re.sub(r'(\!\[.*?\]\(.*?\))\{.*?\}', r'\1', content, flags=re.DOTALL)
                f.seek(0)
                f.write(content)
                f.truncate()
            
            output_widget.insert(tk.END, "Markdown cleanup complete.\n")

            if post_process:
                post_process_with_ai(output_file_path, **ai_params, output_widget=output_widget)

            output_widget.insert(tk.END, f"\n--- Conversion Completed Successfully! ---\n")
        else:
            output_widget.insert(tk.END, f"\n--- Pandoc failed with exit code {process.returncode}. Check logs above. ---\n")

    except Exception as e:
        import traceback
        output_widget.insert(tk.END, f"\n--- AN UNEXPECTED ERROR OCCURRED ---\n{traceback.format_exc()}\n")

def start_conversion(vars_dict):
    """Gathers all inputs from the UI and starts the conversion process."""
    if not all(v.get() for k, v in vars_dict.items() if k in ['pandoc_path', 'file_path', 'output_dir', 'wiki_root_path', 'attachments_dir']):
        messagebox.showerror("Error", "All path and folder fields must be filled out.")
        return

    post_process = vars_dict['post_process_var'].get()
    ai_params = {"gpt_endpoint": vars_dict['gpt_endpoint'].get(), "gpt_key": vars_dict['gpt_key'].get(), "gpt_deployment": vars_dict['gpt_deployment'].get(), "prompt": vars_dict['prompt_text'].get("1.0", "end-1c")}

    if post_process and not all(ai_params.values()):
        messagebox.showerror("Error", "To use AI post-processing, all Azure OpenAI fields and the prompt must be filled out.")
        return

    vars_dict['output_widget'].delete('1.0', tk.END)
    
    thread = threading.Thread(
        target=run_conversion_logic,
        args=(
            vars_dict['pandoc_path'].get(), vars_dict['file_path'].get(),
            vars_dict['output_dir'].get(), vars_dict['attachments_dir'].get(),
            vars_dict['wiki_root_path'].get(), post_process, ai_params,
            vars_dict['output_widget']
        )
    )
    thread.start()

def create_gui():
    """Creates and runs the Tkinter GUI."""
    root = tk.Tk()
    root.title("Pandoc Converter for DOCX")
    root.geometry("850x700")

    main_frame = tk.Frame(root, padx=10, pady=10)
    main_frame.pack(fill=tk.BOTH, expand=True)

    vars_dict = {
        'pandoc_path': tk.StringVar(), 'file_path': tk.StringVar(),
        'output_dir': tk.StringVar(), 'wiki_root_path': tk.StringVar(),
        'attachments_dir': tk.StringVar(value=".attachments"),
        'post_process_var': tk.BooleanVar(), 'gpt_endpoint': tk.StringVar(),
        'gpt_key': tk.StringVar(), 'gpt_deployment': tk.StringVar(),
    }

    settings_frame = tk.LabelFrame(main_frame, text="Settings", padx=10, pady=10)
    settings_frame.pack(fill=tk.X, pady=5)
    tk.Label(settings_frame, text="Pandoc Path:").grid(row=0, column=0, sticky="w", padx=5, pady=3)
    tk.Entry(settings_frame, textvariable=vars_dict['pandoc_path']).grid(row=0, column=1, sticky="ew")
    tk.Button(settings_frame, text="Browse...", command=lambda: vars_dict['pandoc_path'].set(filedialog.askopenfilename(title="Select pandoc.exe"))).grid(row=0, column=2, padx=5)
    settings_frame.columnconfigure(1, weight=1)

    fields_frame = tk.LabelFrame(main_frame, text="File and Folder Selection", padx=10, pady=10)
    fields_frame.pack(fill=tk.X, pady=5)
    tk.Label(fields_frame, text="Document File (DOCX, etc.):").grid(row=0, column=0, sticky="w", padx=5, pady=3)
    tk.Entry(fields_frame, textvariable=vars_dict['file_path']).grid(row=0, column=1, sticky="ew")
    tk.Button(fields_frame, text="Browse...", command=lambda: vars_dict['file_path'].set(filedialog.askopenfilename())).grid(row=0, column=2, padx=5)
    tk.Label(fields_frame, text="Output Folder:").grid(row=1, column=0, sticky="w", padx=5, pady=3)
    tk.Entry(fields_frame, textvariable=vars_dict['output_dir']).grid(row=1, column=1, sticky="ew")
    tk.Button(fields_frame, text="Browse...", command=lambda: vars_dict['output_dir'].set(filedialog.askdirectory())).grid(row=1, column=2, padx=5)
    tk.Label(fields_frame, text="Wiki Root Folder:").grid(row=2, column=0, sticky="w", padx=5, pady=3)
    tk.Entry(fields_frame, textvariable=vars_dict['wiki_root_path']).grid(row=2, column=1, sticky="ew")
    tk.Button(fields_frame, text="Browse...", command=lambda: vars_dict['wiki_root_path'].set(filedialog.askdirectory())).grid(row=2, column=2, padx=5)
    tk.Label(fields_frame, text="Attachments Folder Name:").grid(row=3, column=0, sticky="w", padx=5, pady=3)
    tk.Entry(fields_frame, textvariable=vars_dict['attachments_dir']).grid(row=3, column=1, sticky="ew")
    fields_frame.columnconfigure(1, weight=1)

    ai_frame = tk.LabelFrame(main_frame, text="AI Post-Processing", padx=10, pady=10)
    ai_frame.pack(fill=tk.X, pady=5)
    tk.Checkbutton(ai_frame, text="Post-process with AI", variable=vars_dict['post_process_var']).grid(row=0, column=0, columnspan=2, sticky="w")
    prompt_frame = tk.LabelFrame(ai_frame, text="Prompt", padx=5, pady=5)
    prompt_frame.grid(row=1, column=0, columnspan=2, sticky="ew", padx=5, pady=5)
    prompt_text = scrolledtext.ScrolledText(prompt_frame, wrap=tk.WORD, height=4)
    prompt_text.pack(fill=tk.X, expand=True)
    vars_dict['prompt_text'] = prompt_text
    tk.Label(ai_frame, text="Azure OpenAI Endpoint:").grid(row=2, column=0, sticky="w", padx=5, pady=2)
    tk.Entry(ai_frame, textvariable=vars_dict['gpt_endpoint']).grid(row=2, column=1, sticky="ew", padx=5, pady=2)
    tk.Label(ai_frame, text="Key:").grid(row=3, column=0, sticky="w", padx=5, pady=2)
    tk.Entry(ai_frame, textvariable=vars_dict['gpt_key'], show="*").grid(row=3, column=1, sticky="ew", padx=5, pady=2)
    tk.Label(ai_frame, text="Deployment:").grid(row=4, column=0, sticky="w", padx=5, pady=2)
    tk.Entry(ai_frame, textvariable=vars_dict['gpt_deployment']).grid(row=4, column=1, sticky="ew", padx=5, pady=2)
    ai_frame.columnconfigure(1, weight=1)
    
    output_frame = tk.LabelFrame(main_frame, text="Conversion Log", padx=10, pady=10)
    output_frame.pack(fill=tk.BOTH, expand=True, pady=10)
    output_widget = scrolledtext.ScrolledText(output_frame, wrap=tk.WORD, height=10, bg="#2b2b2b", fg="#d3d3d3")
    output_widget.pack(fill=tk.BOTH, expand=True)
    vars_dict['output_widget'] = output_widget
    
    button_frame = tk.Frame(main_frame)
    button_frame.pack(fill=tk.X, pady=5)
    save_button = tk.Button(button_frame, text="Save Settings", command=lambda: save_config(vars_dict))
    save_button.pack(side=tk.LEFT, expand=True, fill=tk.BOTH, padx=(0,5))
    convert_button = tk.Button(button_frame, text="Convert to Markdown", command=lambda: start_conversion(vars_dict), bg="#0078D4", fg="white", font=("sans-serif", 12, "bold"))
    convert_button.pack(side=tk.LEFT, expand=True, fill=tk.BOTH, padx=(5,0))

    load_config(vars_dict)
    root.mainloop()

def check_dependencies():
    """Checks for required packages before starting the GUI."""
    missing = []
    try: from cryptography.fernet import Fernet
    except ImportError: missing.append("'cryptography'")
    try: from openai import AzureOpenAI
    except ImportError: missing.append("'openai'")
    
    if missing:
        return (f"The following required package(s) are not installed: {', '.join(missing)}\n\n"
                f"Please activate your virtual environment and run:\npip install {' '.join(missing)}")
    return None

if __name__ == "__main__":
    error_message = check_dependencies()
    if error_message:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Missing Dependencies", error_message)
        sys.exit(1)
    
    create_gui()
