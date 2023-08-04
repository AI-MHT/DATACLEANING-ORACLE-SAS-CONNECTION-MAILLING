import cx_Oracle
import saspy
# Oracle database connection details
oracle_username = 'YOUR_ORACLE_USERNAME'
oracle_password = 'YOUR_ORACLE_PASSWORD'
oracle_host = 'YOUR_ORACLE_HOST'
oracle_port = 'YOUR_ORACLE_PORT'
oracle_service_name = 'YOUR_ORACLE_SERVICE_NAME'

# SAS connection details
sas_username = 'YOUR_SAS_USERNAME'
sas_password = 'YOUR_SAS_PASSWORD'

# Extracting database from Oracle and saving it to a local folder
def extract_database():
    # Oracle connection string
    dsn = cx_Oracle.makedsn(oracle_host, oracle_port, service_name=oracle_service_name)
    conn = cx_Oracle.connect(user=oracle_username, password=oracle_password, dsn=dsn)
    cursor = conn.cursor()

    # Execute the SQL query to extract the SINISTRE database
    sql_query = 'SELECT * FROM SINISTRE'
    cursor.execute(sql_query)
    result = cursor.fetchall()

    # Save the result to a local file
    file_path = 'SINISTRE_RPA/SINISTRE.csv'
    with open(file_path, 'w') as file:
        for row in result:
            file.write(','.join(map(str, row)) + '\n')

    cursor.close()
    conn.close()

    print('Database extracted and saved to', file_path)

# Uploading file to SAS
def upload_to_sas():
    sas = saspy.SASsession(username=sas_username, password=sas_password)

    # Upload the file to SAS using proc import
    sas.submit(
        f'''
        proc import datafile='/path/to/SINISTRE_RPA/SINISTRE.csv'
            out=work.SINISTRE
            dbms=csv
            replace;
            /* Define your columns here based on the database structure */
            /* Example: */
            /* getnames=yes; */
            /* guessingrows=1000; */
        run;
        '''
    )

    sas.disconnect()

    print('File uploaded to SAS')

# Main script
if __name__ == '__main__':
    extract_database()
    upload_to_sas()
