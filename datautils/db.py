import pyodbc
from sqlalchemy import create_engine
import os
from sqlalchemy import text

class MachineMSSQLServer:
    def __init__(self, server, database, username, password):
        self.engine =  create_engine('mssql+pyodbc://'+username+':'+password+'@'+server+'/'+database+'?driver=ODBC+Driver+18+for+SQL+Server&TrustServerCertificate=yes')

    def save_id_to_file(self,filename, id):
        with open(filename, 'w') as f:
            f.write(str(id))

    def read_id_from_file(self,filename):
        if os.path.isfile(filename):
            with open(filename, 'r') as f:
                return int(f.read().strip())
        else:
            return 0

    def load_data(self):
        with self.engine.connect() as connection:
            q = text(f"SELECT '_________' as RFID,* FROM dbo.MachineIntegration where ID > {self.read_id_from_file('id.txt')} order by ID ASC")
            
            return connection.execute(q).fetchall()

    def insert_data(self,rfid,bcode):
        query = text(f"""INSERT INTO [ActiveSooperWizerNCL].[Essentials].[Tag] 
                (TagID, GroupID) 
                VALUES 
                ('{rfid}',{bcode})""")
        print(query)
        with self.engine.connect() as connection:
            _res = connection.execute(query)
            connection.commit()
            if _res.rowcount > 0:
                return 0
            else:
                return 1
    
    def upload_data(self,rfid,bcode,bnd,ext):
        query2 = text(f"""UPDATE [ActiveSooperWizerNCL].[Essentials].[Tag] 
                SET GroupID = {bcode},
                BundleID = {bnd}
                where TagID = '{rfid}'
                """)
        query = text(f"""MERGE [ActiveSooperWizerNCL].[Essentials].[Tag] AS target
                    USING (SELECT '{rfid}' as TagID, {bcode} as GroupID, {bnd} as BundleID, {ext} as extInfo) AS source
                    ON (target.TagID = source.TagID)
                    WHEN MATCHED THEN 
                        UPDATE SET target.GroupID = source.GroupID, target.BundleID = source.BundleID,target.UpdatedAt = GETDATE(),target.extInfo = source.extInfo
                    WHEN NOT MATCHED THEN
                    INSERT (TagID, GroupID,BundleID,extInfo) 
                        VALUES (source.TagID, source.GroupID,source.BundleID,source.extInfo);""")
        #print(query)
        try:
            with self.engine.connect() as connection:
                _res = connection.execute(query)
                connection.commit()
                if _res.rowcount > 0:
                    return 0
                else:
                    return 1
        except Exception as ex:
            print(str(ex))
            return 1

# _machinedb = MachineMSSQLServer('172.16.20.1', 'ActiveSooperWizerNCL', 'sa', 'wimetrix')
# _lbldata = _machinedb.load_data()

# for row in _lbldata:
#     _lbl = row._asdict()
#     print(_lbl)
#     # if _machinedb.upload_data() < 1:
#     #     print("Couldnt Upload Data Check Network Connection !")
#     _machinedb.save_id_to_file('id.txt',_lbl['ID'])
