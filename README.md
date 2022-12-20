# excel2kml
This tool generates a .kml file from a given .csv or .xlsx file.

# Installation:
Python version >= 3.8.5 is required.  
`pip install -r requirements.txt`

# Usage:
Place inside working directory a .csv or .xlsx file with id | lat | lon columns.  
The first row must be the header with the exact names: id, lat, lon.  
Creates a directory <datetime>_KML where you can find the exported .kml file

# File examples:
        Points:
            id	    lat	        lon
            point1	38.25421	23.28931
            point2	38.27523	23.31792
            point3	38.25991	23.36405
            point4	38.23862	23.32512
            point5	38.25411	23.28921
            point6	38.27513	23.31782
            point7	38.25981	23.36395
            point8	38.23852	23.32502
        this will create a kml with 8 points

        Lines:
            id	    lat	        lon
            line1	38.25421	23.28931
                    38.27523	23.31792
                    38.25991	23.36405
            line2	38.23862	23.32512
                    38.25411	23.28921
            line3	38.27513	23.31782
                    38.25981	23.36395
                    38.23852	23.32502
            this will create a kml with 3 lines.
            Lines 1 and 3 will have 3 vertices and line2 will have 2 vertices.

        Polygons:
            id	    lat	        lon
            poly1	38.25421	23.28931
                    38.27523	23.31792
                    38.25991	23.36405
                    38.23862	23.32512
            poly2	38.25411	23.28921
                    38.27513	23.31782
                    38.25981	23.36395
            this will create a kml with 2 polygons.
            poly1 will have 4 vertices and poly2 will have 3 vertices.

# Arguments: 
`<kml type> <import_filename.(csv or xlsx)> <output_kml_filename.kml>`

Where `<kml type>` must be point, line or polygon.

# Execution (example): 
`python excel2kml.py line inputlines.xlsx outputlines.kml`
    
# Help:
`python excel2kml.py -h`
