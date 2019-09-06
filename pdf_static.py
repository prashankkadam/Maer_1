import pandas as pd
from fpdf import FPDF


def add_image(image_path, dataframe):
    pdf = FPDF()
    pdf.add_page()
    pdf.image(image_path, x=7.5, y=6, w=75)
    pdf.set_font("Arial", size=12)
    pdf.ln(20)  # move 20 down
    # pdf.cell(200, 10, txt="{}".format(image_path), ln=1)
    font = "Arial"
    font_size = 14
    txt = "{}".format(font, font_size)
    pdf.cell(0, 10, txt=txt, ln=1)

    pdf.ln(10)
    dataframe = dataframe.loc[:, 'Vessel':'Type']

    data = dataframe.values.tolist()
    pdf.set_font("Arial", size=12)
    spacing = 1
    col_width = pdf.w / 4.5
    row_height = 15
    for row in data:
        for item in row:
            pdf.cell(col_width, row_height * spacing,
                     txt=str(item), border=1)
        pdf.ln(row_height * spacing)

    pdf.output("add_image.pdf")


if __name__ == '__main__':
    df_static = pd.read_excel('Static_data_final.xlsx')
    vesselName = input("Enter the vessel name: ")
    df_vessel_data = df_static[df_static['Vessel'] == vesselName]
    df_vessel_data = df_vessel_data.fillna(0)
    add_image('maersk_logo_1.jpg', df_vessel_data)
