import streamlit as st
import csv
from xml.etree.ElementTree import Element, SubElement, tostring
from xml.dom.minidom import parseString
from datetime import datetime
import time

# Hardcoded mapping
mapping = {
    "Receipt Bank Code": "BankCode",
    "Name On Card": "CreditCardName",
    "Masked Credit Card No.": "CreditCardNumber",
    "Credit Card Type": "CreditCardType",
    "Bank Deposit Date": "Date",
    "ID": "ID",
    "Receipt No.": "ReceiptNumber",
    "Total Amount Charged": "TotalAmount",
    "Banked Amount": "GLAmount",
    "GL Code": "GLCode",
    "GL Description": "GLDescription",
    "GL Receipt Name": "GLReceiptName"
}

# Function to create the XML from CSV
def create_xml(file):
    try:
        # Root element
        root = Element('Synergetic')
        root.set("FileCreatedDate", datetime.now().strftime("%d/%m/%Y"))
        root.set("FileVersion", "1.00")
        root.set("SupplierReference", "KINGS")

        # Read CSV file
        csv_reader = csv.DictReader(file.decode("utf-8").splitlines())
        for row in csv_reader:
            # Create <Receipt> element
            receipt = SubElement(root, "Receipt")
            for csv_field, xml_field in mapping.items():
                if xml_field in ["GLAmount", "GLCode", "GLDescription", "GLReceiptName"]:
                    # Skip GLPayment attributes for now
                    continue
                receipt.set(xml_field, row.get(csv_field, ""))

            # Add <GLPayment> element inside <Receipt>
            gl_payment = SubElement(receipt, "GLPayment")
            for csv_field, xml_field in mapping.items():
                if xml_field in ["GLAmount", "GLCode", "GLDescription", "GLReceiptName"]:
                    gl_payment.set(xml_field, row.get(csv_field, ""))
        
        # Convert to pretty XML
        xml_string = tostring(root, encoding='unicode')
        dom = parseString(xml_string)
        return dom.toprettyxml(indent="  ")
    except Exception as e:
        return f"Error: {e}"

# Streamlit App
def main():
    st.title("CSV to XML Converter")
    st.write("Upload a CSV file and download the converted XML.")

    # Upload CSV file
    csv_file = st.file_uploader("Upload CSV File", type=["csv"])

    if not csv_file:
        st.info("Please upload a CSV file to begin.")
        return

    # Convert button
    if st.button("Convert to XML"):
        with st.spinner("Converting..."):
            time.sleep(3)
            try:
                # Perform conversion
                xml_output = create_xml(csv_file.getvalue())
                if xml_output.startswith("Error:"):
                    st.error(xml_output)
                else:
                    st.success("Conversion successful!")
                    st.download_button(
                        label="Download XML",
                        data=xml_output,
                        file_name="output.xml",
                        mime="application/xml",
                        type="secondary",
                        icon=":material/download:"
                    )
                    
            except Exception as e:
                st.error(f"An error occurred: {e}")

if __name__ == "__main__":
    main()
