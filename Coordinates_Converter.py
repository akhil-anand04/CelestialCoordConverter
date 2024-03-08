import openpyxl
import astropy.units as u
import astropy.coordinates as coord

def extract_data_from_excel(file_path):
    workbook = openpyxl.load_workbook(file_path)
    worksheet = workbook.active
    table_data = {}
    for row in worksheet.iter_rows(min_row=2, values_only=True):
        target_name = row[0]
        ra_str = row[1]
        dec_str = row[2]
        if ra_str is None or dec_str is None:
            continue
        try:
            ra_decimal = convert_ra_to_decimal(ra_str)
            dec_decimal = convert_dec_to_decimal(dec_str)
            table_data[target_name] = {'R.A. (J2000)': ra_decimal, 'Decl.': dec_decimal}
        except (ValueError, IndexError):
            continue
    workbook.close()
    return table_data

def convert_ra_to_decimal(ra_str):
    ra_components = ra_str.split()
    ra_decimal = (float(ra_components[0]) + float(ra_components[1])/60 + float(ra_components[2])/3600) * 15
    return ra_decimal

def convert_dec_to_decimal(dec_str):
    dec_components = dec_str.split()
    sign = 1 if dec_components[0][0] == '+' else -1
    dec_decimal = sign * ((float(dec_components[0])) + float(dec_components[1])/60 + float(dec_components[2])/3600)
    return dec_decimal

def compute_table(table_data):
    ra = [entry['R.A. (J2000)'] for entry in table_data.values()]
    dec = [entry['Decl.'] for entry in table_data.values()]
    target_names = list(table_data.keys())
    return ra, dec, target_names

def galactic(ra, dec, target_names):
    galactic_coords = []
    for i in range(len(ra)):
        c = coord.SkyCoord(ra[i], dec[i], unit=(u.deg, u.deg))
        galactic_coord = c.transform_to(coord.Galactic())
        galactic_coords.append((target_names[i], galactic_coord))
    return galactic_coords

def filter_galactic(galactic_coords):
    filtered_coords = []
    for target_name, coord in galactic_coords:
        l = coord.l.deg
        if 0 < l < 180:
            filtered_coords.append((target_name, coord))
    return filtered_coords

def export_to_excel(filtered_coords):
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = "Filtered Galactic Coordinates"
    worksheet['A1'] = "Target Name"
    worksheet['B1'] = "Galactic Longitude (l)"
    worksheet['C1'] = "Galactic Latitude (b)"
    for i, (target_name, coord) in enumerate(filtered_coords, start=2):
        worksheet[f'A{i}'] = target_name
        worksheet[f'B{i}'] = coord.l.deg
        worksheet[f'C{i}'] = coord.b.deg
    workbook.save(filename="Filtered_Galactic_Coordinates.xlsx")

def main():
    file_path = 'Coords.xlsx'
    table_data = extract_data_from_excel(file_path)
    ra, dec, target_names = compute_table(table_data)
    galactic_coords = galactic(ra, dec, target_names)
    filtered_coords = filter_galactic(galactic_coords)
    if filtered_coords:
        export_to_excel(filtered_coords)
    else:
        print("No filtered coordinates to export.")

main()
