import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload
import threading
import os
 
 
# Initialize global variables
creds = None
CLIENT_SECRET_FILE = ''  # Path to your credentials JSON file
SCOPES = ['https://www.googleapis.com/auth/drive']  # Full access to Google Drive
service = None
items = []
file_tree = None
credentials_label = None  # Declare the label as a global variable
 
 
 
 
def authenticate():
   global creds
   if os.path.exists('token.json'):
       creds = Credentials.from_authorized_user_file('token.json', SCOPES)
       if not creds.valid:
           try:
               creds.refresh(Request())
           except Exception as e:
               os.remove('token.json')
               flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET_FILE, SCOPES)
               creds = flow.run_local_server(port=0)
               with open('token.json', 'w') as token:
                   token.write(creds.to_json())
   else:
       flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET_FILE, SCOPES)
       creds = flow.run_local_server(port=0)
       with open('token.json', 'w') as token:
           token.write(creds.to_json())
   return creds
 
 
 
 
def list_files():
   global service, items
   creds = authenticate()
   service = build('drive', 'v3', credentials=creds)
   results = service.files().list(
       pageSize=100,
       fields="nextPageToken, files(id, name, mimeType)",
       includeItemsFromAllDrives=True,
       supportsAllDrives=True
   ).execute()
   items = results.get('files', [])
   file_tree.delete(*file_tree.get_children())
   if not items:
       file_tree.insert("", 'end', text="No files found.")
   else:
       for item in items:
           file_tree.insert("", 'end', values=(item['name'], item['id'], item['mimeType']))
 
 
 
 
def upload_file():
   file_path = filedialog.askopenfilename()
   if file_path:
       threading.Thread(target=upload_thread, args=(file_path,)).start()
 
 
 
 
def upload_thread(file_path):
   creds = authenticate()
   service = build('drive', 'v3', credentials=creds)
   file_name = os.path.basename(file_path)
   file_metadata = {'name': file_name}
   media = MediaFileUpload(file_path, resumable=True)
   file = service.files().create(body=file_metadata, media_body=media, fields='id').execute()
   print(f'File uploaded: {file.get("id")}')
 
 
 
 
def download_file():
   selected_item = file_tree.selection()
   if selected_item:
       threading.Thread(target=download_thread, args=(selected_item,)).start()
 
 
 
 
def download_thread(selected_item):
   global items
   try:
       creds = authenticate()
       service = build('drive', 'v3', credentials=creds)
       file_id = file_tree.item(selected_item, 'values')[1]
       file_metadata = service.files().get(fileId=file_id, fields='mimeType, name').execute()
       mime_type = file_metadata['mimeType']
       file_name = file_tree.item(selected_item, 'values')[0]
 
 
       # Handling different file types, including Google Slides
       if 'google-apps' in mime_type:
           export_types = {
               'application/vnd.google-apps.document': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
               'application/vnd.google-apps.spreadsheet': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
               'application/vnd.google-apps.presentation': 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
           }
           if mime_type in export_types:
               request = service.files().export_media(fileId=file_id, mimeType=export_types[mime_type])
               file_name += '.pptx' if 'presentation' in mime_type else '.docx'
           else:
               raise Exception(f"File type '{mime_type}' is not exportable.")
       else:
           request = service.files().get_media(fileId=file_id)
 
 
       download_directory = filedialog.askdirectory()  # Prompt user to select download directory
       if download_directory:
           file_path = os.path.join(download_directory, file_name)
           with open(file_path, 'wb') as fh:
               downloader = MediaIoBaseDownload(fh, request)
               done = False
               while not done:
                   status, done = downloader.next_chunk()
                   print(f'Download {int(status.progress() * 100)}% - {file_name}')
           messagebox.showinfo("Download", f"File downloaded successfully! {file_path}")
   except Exception as e:
       print(f"Error downloading file: {e}")
       messagebox.showerror("Error", f"Error downloading file: {e}")
 
 
 
 
def select_credentials_file():
   global CLIENT_SECRET_FILE, credentials_label
   file_path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
   if file_path:
       CLIENT_SECRET_FILE = file_path
       credentials_label.config(text=f"Credentials file: {CLIENT_SECRET_FILE}")
 
 
 
 
def setup_gui(root):
   global file_tree, credentials_label
   root.title("Google Drive File Manager - The Pycodes")
   root.geometry("700x450")
 
 
   frame = ttk.Frame(root)  # Frame to hold the Treeview and Scrollbar
   frame.pack(expand=True, fill='both', padx=10, pady=10)
 
 
   # Setup Treeview
   cols = ('File Name', 'File ID', 'MIME Type')
   file_tree = ttk.Treeview(frame, columns=cols, show='headings')
   for col in cols:
       file_tree.heading(col, text=col)
   file_tree.pack(side=tk.LEFT, expand=True, fill='both')
 
 
   # Setup Scrollbar
   scrollbar = ttk.Scrollbar(frame, orient="vertical", command=file_tree.yview)
   scrollbar.pack(side=tk.RIGHT, fill='y')
   file_tree.configure(yscrollcommand=scrollbar.set)
 
 
   # Setup Buttons
   ttk.Button(root, text="List Files", command=list_files).pack(side=tk.TOP, pady=5)
   ttk.Button(root, text="Upload File", command=upload_file).pack(side=tk.TOP, pady=5)
   ttk.Button(root, text="Download Selected File", command=download_file).pack(side=tk.TOP, pady=5)
   ttk.Button(root, text="Select Credentials File", command=select_credentials_file).pack(side=tk.TOP, pady=5)
 
 
   # Display selected credentials file
   credentials_label = ttk.Label(root, text="Credentials file: ")
   credentials_label.pack(side=tk.TOP, pady=5)
 
 
 
 
root = tk.Tk()
setup_gui(root)
root.mainloop()
