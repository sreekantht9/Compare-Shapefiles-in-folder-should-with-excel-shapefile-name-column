import os
import geopandas as gpd
import pandas as pd

def list_shapefile_info(input_folder, output_file):
    # Initialize the output DataFrame
    output_df = pd.DataFrame(columns=['S.NO', 'ExcelFileName', 'ExcelGeometryType',
                                      'Shapefile_Folder', 'Shapefile_GeometryType', 'MatchStatus'])
    
    # Get a list of all files in the input folder
    files = os.listdir(input_folder)
    
    # Filter only shapefiles from the list of files
    shapefile_list = [file for file in files if file.lower().endswith('.shp')]
    
    # Loop through each shapefile
    for idx, shapefile in enumerate(shapefile_list, start=1):
        # Get the full path of the shapefile
        shapefile_path = os.path.join(input_folder, shapefile)
        
        # Read the shapefile using geopandas
        gdf = gpd.read_file(shapefile_path)
        
        # Get the Excel file name without the extension
        excel_file_name = os.path.splitext(shapefile)[0]
        
        # Get the geometry type of the shapefile
        shapefile_geometry_type = gdf.geom_type.unique()[0]
        
        # Append the information to the output DataFrame
        output_df = output_df.append({
            'S.NO': idx,
            'ExcelFileName': excel_file_name,
            'ExcelGeometryType': 'N/A',  # Replace 'N/A' with the actual Excel Geometry Type if available
            'Shapefile_Folder': input_folder,
            'Shapefile_GeometryType': shapefile_geometry_type,
            'MatchStatus': 'N/A'  # Replace 'N/A' with the actual Match Status if available
        }, ignore_index=True)
    
    # Write the DataFrame to the output Excel file
    output_df.to_excel(output_file, index=False)

if __name__ == "__main__":
    # Replace 'input_folder' and 'output_file' with the appropriate folder and file paths
    
    input_folder = "D:/Projects/Warangal/Created_Shapefiles/Check/"
    output_file = "D:/Projects/Warangal/Created_Shapefiles/Check/0.Geometry_Check.xlsx"
    
    list_shapefile_info(input_folder, output_file)

# Read the output Excel file into a DataFrame
result_df = pd.read_excel(output_file)

# Access the desired columns as lists
shapefile_names = result_df['ExcelFileName'].tolist()
geometry_types = result_df['Shapefile_GeometryType'].tolist()

print("Shapefile names:", shapefile_names)
print("Geometry types:", geometry_types)

with pd.ExcelWriter(output_file) as writer:
    output_df.to_excel(writer, index=False)