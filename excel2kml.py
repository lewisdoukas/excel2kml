import traceback, sys, os, datetime
import simplekml
import pandas as pd


banner = \
'''
 _______   _______  _____ _     _        _   _____  ___ _     
|  ___\ \ / /  __ \|  ___| |   | |      | | / /|  \/  || |    
| |__  \ V /| /  \/| |__ | |   | |_ ___ | |/ / | .  . || |    
|  __| /   \| |    |  __|| |   | __/ _ \|    \ | |\/| || |    
| |___/ /^\ \ \__/\| |___| |___| || (_) | |\  \| |  | || |____
\____/\/   \/\____/\____/\_____/\__\___/\_| \_/\_|  |_/\_____/
                                                              
                                               by Ilias Doukas
'''


def help():
    h = """

    This tool generates a .kml file from a given .csv or .xlsx file.

    Usage:
        Place inside working directory a .csv or .xlsx file with id | lat | lon columns.
        The first row must be the header with the exact names: id, lat, lon.
        Creates a directory <datetime>_KML where you can find the exported .kml file

    File examples:
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

    Arguments: 
        <kml type> <import_filename.(csv or xlsx)> <output_kml_filename.kml>

        Where <kml type> must be point, line or polygon

    Execution (example): 
        python excel2kml.py line inputlines.xlsx outputlines.kml
    
    Help:
        python excel2kml.py -h

    """

    print(h)


class KmlCreator():
    def __init__(self):
        self.kml = simplekml.Kml()
        self.__create_export_dir()
    
    def create_kml(self, kml_type, filename, exp_filename):
        try:
            df = self.__read_data(filename)

            if df.empty: 
                print("Imported file is empty. Try another one.")
                return
            
            if kml_type == "point":
                df.apply(self.__zip_points, axis= 1)
            elif kml_type == "line":
                self.__zip_lines(df)
            else:
                self.__zip_polygons(df)

            self.__save_kml(f"{self.export_dir}/{exp_filename}")
            print(f"\n✓ {kml_type.capitalize()} KML file {exp_filename} has been successfully created!")
        except Exception as e:
            ex = "An unexpected error occured while generating point kml. Check log file\n"
            error = f"{ex}{traceback.format_exc()}\n{e}"
            print(ex)
            self.__write_log(error)


    def __create_export_dir(self):
        now = datetime.datetime.now().strftime("%Y.%m.%d")

        if getattr(sys, "frozen", False):
            currentDirName = os.path.dirname(sys.executable).replace('\\', '/')
        elif __file__:
            currentDirName = os.path.dirname(__file__).replace('\\', '/')

        if not currentDirName:
            currentDirName = os.path.abspath(os.getcwd())

        self.export_dir = f"{currentDirName}/{now}_KML"
        dir_exists = os.path.isdir(self.export_dir)

        if not dir_exists:
            os.mkdir(self.export_dir)

    def __write_log(self, error):
        now = datetime.datetime.now().strftime("%Y.%m.%d_%H.%M.%S")
        with open(f"{self.export_dir}/{now}_error.log", "a") as file:
            file.write(error)

    def __read_data(self, filename):
        splittedFname = filename.split('/')
        fName = splittedFname[-1].split('.')
        fType = fName[-1]
        df = pd.DataFrame()

        if fType == "csv":
            df = pd.read_csv(filename)
        else:
            df = pd.read_excel(filename)

        return(df)
    
    def __save_kml(self, filename):
        self.kml.save(filename)

    def __zip_points(self, pointDf):
        self.kml.newpoint(name= pointDf['id'], coords= [(pointDf['lon'], pointDf['lat'])])
    
    def __zip_lines(self, lineDf):
        for index, row in lineDf.iterrows():
            if not pd.isna(row['id']):
                line = self.kml.newlinestring(name= row['id'], coords= [(row['lon'], row['lat'])])
                lineCoords = [(row['lon'], row['lat'])]
            else:
                lineCoords.append((row['lon'], row['lat']))
                line.coords = lineCoords
    
    def __zip_polygons(self, polyDf):
        for index, row in polyDf.iterrows():
            if not pd.isna(row['id']):
                poly = self.kml.newpolygon(name= row['id'], outerboundaryis= [(row['lon'], row['lat'])])
                polyCoords = [(row['lon'], row['lat'])]
            else:
                polyCoords.append((row['lon'], row['lat']))
                poly.outerboundaryis = polyCoords
                

def main():
    print(banner)

    if len(sys.argv) > 1:
        kml_type = str(sys.argv[1])

        if kml_type == "-h":
            help()
            return

        if kml_type in ["point", "line", "polygon"]:

            if len(sys.argv) > 2:
                imp_filename = str(sys.argv[2])

                splittedFname = imp_filename.split('.')
                fName = splittedFname[-1].split('.')
                fType = fName[-1]

                if len(splittedFname) > 1 and (fType == "csv" or fType == "xlsx"):
                    if len(sys.argv) > 3:
                        exp_filename = str(sys.argv[3]) if ".kml" in str(sys.argv[3]) else str(sys.argv[3]) + ".kml"
                    else:
                        exp_filename = f"output_{kml_type}.kml"
                
                    kml = KmlCreator()
                    kml.create_kml(kml_type, imp_filename, exp_filename)
                else:
                    print("✕ Please select a valid .csv or .xlsx file.\n")
            else:
                print("✕ Please input a valid .csv or .xlsx filename.\n")
        else:
            print("✕ Supported kml types are point, line, polygon.\n")
    else:
        print("✕ Please type:\n\npython excel2kml.py -h\n\nfor help.")


if __name__ == "__main__":
    main()
