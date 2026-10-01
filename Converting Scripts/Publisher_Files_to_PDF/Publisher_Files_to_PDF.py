import os
import win32com.client

folder = r"path/to/publisher_documents"

publisher = win32com.client.Dispatch("Publisher.Application")

for root, dirs, files in os.walk(folder):
    for filename in files:
        if filename.lower().endswith(".pub"):
            pub = os.path.join(root, filename)
            pdf = os.path.splitext(pub)[0] + ".pdf"

            print("Converting:", pub)

            doc = publisher.Open(pub)
            doc.ExportAsFixedFormat(2, pdf)
            doc.Close()

            print("Created:", pdf)

publisher.Quit()

print("Finished!")
input("Press Enter to close...")
