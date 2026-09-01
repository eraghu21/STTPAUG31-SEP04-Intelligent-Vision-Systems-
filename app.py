import streamlit as st
import pandas as pd
from fpdf import FPDF
import re
import pyAesCrypt
import requests
import os
import base64

# ====================== CONFIG ======================
st.set_page_config(page_title="SVCE STTP Certificate Generator", layout="centered")

# ====================== COUNTERS ======================
def update_visit_count():
    count_file = "counter.txt"
    if not os.path.exists(count_file):
        with open(count_file, "w") as f:
            f.write("0")
    with open(count_file, "r") as f:
        count = int(f.read())
    if "counted" not in st.session_state:
        count += 1
        with open(count_file, "w") as f:
            f.write(str(count))
        st.session_state.counted = True
    return count

def update_download_count():
    count_file = "downloads.txt"
    if not os.path.exists(count_file):
        with open(count_file, "w") as f:
            f.write("0")
    with open(count_file, "r") as f:
        count = int(f.read())
    count += 1
    with open(count_file, "w") as f:
        f.write(str(count))
    return count

def get_download_count():
    count_file = "downloads.txt"
    if not os.path.exists(count_file):
        return 0
    with open(count_file, "r") as f:
        return int(f.read())

visit_count = update_visit_count()
download_total = get_download_count()

# ====================== HEADER ======================
st.title("SVCE_CSE: STTP Certificate Generator")
st.markdown(f"<div style='text-align:right; color:gray;'>👁️ Day Visits: {visit_count}</div>", unsafe_allow_html=True)
st.markdown(f"<div style='text-align:right; color:gray;'>📥 Day Downloads: {download_total}</div>", unsafe_allow_html=True)

```python
# ====================== LOAD & DECRYPT EXCEL ======================
buffer_size = 64 * 1024
password = st.secrets["excel_password"]

encrypted_url = (
    "https://raw.githubusercontent.com/eraghu21/"
    "STTPAUG31-SEP04-Intelligent-Vision-Systems-/main/"
    "registrations.xlsx.aes"
)

enc_file = "registrations.xlsx.aes"
dec_file = "registrations.xlsx"

try:
    resp = requests.get(encrypted_url, timeout=30)

    if resp.status_code != 200 or len(resp.content) == 0:
        st.error("❌ Failed to download participant data.")
        st.stop()

    with open(enc_file, "wb") as f:
        f.write(resp.content)

    pyAesCrypt.decryptFile(
        enc_file,
        dec_file,
        password,
        buffer_size
    )

    df = pd.read_excel(dec_file)

    # Remove temporary files
    if os.path.exists(enc_file):
        os.remove(enc_file)

    if os.path.exists(dec_file):
        os.remove(dec_file)

except Exception:
    st.error(
        "❌ Error loading participant data. "
        "Please try again later."
    )
    st.stop()


# ====================== CLEAN COLUMN NAMES ======================
df.columns = (
    df.columns
    .astype(str)
    .str.replace("\ufeff", "", regex=False)
    .str.strip()
    .str.lower()
    .str.replace(r"\s+", "_", regex=True)
)

# Convert your Excel column names to standard names
column_mapping = {
    "email_id": "email",
    "email_address": "email",
    "mail": "email",

    "college": "college_name",
    "college_name": "college_name",

    "s._no": "s_no",
    "s_no": "s_no",
}

df.rename(columns=column_mapping, inplace=True)


# ====================== VALIDATE REQUIRED COLUMNS ======================
required_columns = [
    "email",
    "name",
    "designation",
    "college_name",
]

missing_columns = [
    column for column in required_columns
    if column not in df.columns
]

if missing_columns:
    st.error(
        f"❌ Required column(s) missing: "
        f"{', '.join(missing_columns)}"
    )

    st.error(
        f"Available columns: "
        f"{', '.join(df.columns.tolist())}"
    )

    st.stop()


# ====================== ATTENDANCE ======================
# Your current Excel file does not contain an attendance column.
# Therefore, set the default attendance to 4.
#
# If you later add an attendance column to Excel,
# the application will automatically use it.

if "attendance" not in df.columns:
    df["attendance"] = 4


# Make attendance numeric and safely handle empty values
df["attendance"] = pd.to_numeric(
    df["attendance"],
    errors="coerce"
).fillna(0)


# ====================== EMAIL INPUT ======================
email_input = st.text_input(
    "📧 Enter your registered Email"
)


def is_valid_email(email):
    pattern = r"^[\w\.-]+@[\w\.-]+\.\w+$"
    return bool(re.match(pattern, email.strip()))


# ====================== GENERATE CERTIFICATE ======================
if st.button("Generate Certificate"):

    if not is_valid_email(email_input):

        st.warning(
            "⚠️ Please enter a valid email address."
        )

    else:

        # Clean email column before comparison
        df["email"] = (
            df["email"]
            .astype(str)
            .str.strip()
            .str.lower()
        )

        entered_email = (
            email_input
            .strip()
            .lower()
        )

        # Find participant
        match = df[
            df["email"] == entered_email
        ]

        if not match.empty:

            row = match.iloc[0]

            # ====================== PARTICIPANT DATA ======================
            attendance = float(row["attendance"])

            name = str(row["name"]).strip()

            designation_raw = str(
                row["designation"]
            ).strip()

            college = str(
                row["college_name"]
            ).strip()


            # ====================== ATTENDANCE CHECK ======================
            if attendance >= 3:

                # ====================== DESIGNATION ======================
                designation_lower = (
                    designation_raw.lower()
                )

                if "assistant" in designation_lower:
                    designation = "Assistant Professor"

                elif "associate" in designation_lower:
                    designation = "Associate Professor"

                elif "professor" in designation_lower:
                    designation = "Professor"

                else:
                    designation = designation_raw.title()


                # ====================== CREATE PDF ======================
                pdf = FPDF(
                    orientation="L",
                    unit="mm",
                    format="A4"
                )

                pdf.add_page()

                # Certificate background
                pdf.image(
                    bg_image_path,
                    x=0,
                    y=0,
                    w=297,
                    h=210
                )

                # Font
                pdf.add_font(
                    "AlexBrush",
                    "",
                    "AlexBrush-Regular.ttf",
                    uni=True
                )


                # ====================== NAME ======================
                pdf.ln(65)

                pdf.set_font(
                    "AlexBrush",
                    "",
                    54
                )

                pdf.set_x(10)

                pdf.set_text_color(
                    212,
                    175,
                    55
                )

                pdf.cell(
                    0,
                    12,
                    txt=name.title(),
                    ln=True,
                    align="C"
                )


                # ====================== DESIGNATION ======================
                pdf.ln(2)

                pdf.set_font(
                    "AlexBrush",
                    "",
                    22
                )

                pdf.set_x(15)

                pdf.set_text_color(
                    0,
                    0,
                    0
                )

                pdf.cell(
                    0,
                    10,
                    designation.title(),
                    ln=True,
                    align="C"
                )


                # ====================== COLLEGE ======================
                pdf.ln(1)

                pdf.set_font(
                    "AlexBrush",
                    "",
                    22
                )

                pdf.cell(
                    0,
                    10,
                    college.title(),
                    ln=True,
                    align="C"
                )


                # ====================== FILE NAME ======================
                safe_name = re.sub(
                    r'[\\/:*?"<>|]+',
                    "",
                    name
                ).strip()

                cert_filename = (
                    f"certificate_"
                    f"{safe_name.replace(' ', '_')}.pdf"
                )


                # ====================== SAVE PDF ======================
                pdf.output(cert_filename)


                # ====================== DOWNLOAD ======================
                with open(
                    cert_filename,
                    "rb"
                ) as f:

                    pdf_data = f.read()


                st.success(
                    "✅ Certificate generated successfully!"
                )

                st.download_button(
                    label="📥 Download Certificate",
                    data=pdf_data,
                    file_name=cert_filename,
                    mime="application/pdf"
                )

                update_download_count()


                # Remove generated PDF after preparing download
                if os.path.exists(cert_filename):
                    os.remove(cert_filename)


            else:

                st.warning(
                    "⚠️ Your attendance is less than required."
                )


        else:

            st.error(
                "❌ Email not found in the registration records."
            )
```

### One important point about attendance

Your **current Excel file does not contain `attendance`**, according to the columns Streamlit detected:

```text
s. no
email id
name
designation
college
```

Therefore, in the code above I have temporarily used:

```python
if "attendance" not in df.columns:
    df["attendance"] = 4
```

That means **every registered participant will currently be treated as having attendance = 4**.

If your actual requirement is that attendance must come from Excel, then you should add an `attendance` column to the encrypted Excel:

```text
S. No | email id | name | designation | college | attendance
1     | lakshman@nbkrist.org | Dr.LAKSHMANA RAO BATTARUSETTY | ASSOCIATE PROFESSOR | NBKR INSTITUTE OF SCIENCE AND TECHNOLOGY | 4
```

Then the code will automatically use the real attendance value.

Your original certificate generation layout—including the background, AlexBrush font, name positioning, designation and college positioning—is retained.

**I recommend adding the `attendance` column rather than hard-coding `4`**, especially if you will have participants with different attendance counts.
