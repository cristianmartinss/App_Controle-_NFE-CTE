import pymssql

#  N F E
def buscarNotas(df):
    conn = pymssql.connect(server='', user='', password='', database='')
    vlist = ','.join(["'%s'" % v for v in df[df.columns[0]].values])
    query = "SELECT F1_CHVNFE, F1_DTDIGIT FROM SF1010 WHERE D_E_L_E_T_ = '' AND F1_CHVNFE IN ({})".format(vlist)
    cursor = conn.cursor()  
    cursor.execute(query)
    row = cursor.fetchall()
    result = {}
    for x in row:
        result[x[0]] = x[1]
    return result

def buscarCnpj_Nfe(df):
    conn = pymssql.connect(server='', user='', password='', database='')
    vlist = ','.join(["'%s'" % v for v in df[df.columns[1]]])
    query = "SELECT A2_CGC, RTRIM(A2_NOME) FROM SA2010 WHERE A2_CGC IN({})".format(vlist)
    cursor = conn.cursor()  
    cursor.execute(query)
    row = cursor.fetchall()
    result = {}
    for x in row:
        result[x[0]] = x[1]
    return result

def buscarcc(df):
    conn = pymssql.connect(server='', user='', password='', database='')
    vlist = ','.join(["'%s'" % v for v in df[df.columns[1]]])
    query = "WITH CTE AS ( SELECT SA2010.A2_CGC, SD1010.D1_CC, ROW_NUMBER() OVER (PARTITION BY SA2010.A2_CGC ORDER BY SD1010.D1_EMISSAO DESC) AS rn FROM SD1010 INNER JOIN SA2010 ON SA2010.A2_COD = SD1010.D1_FORNECE WHERE SA2010.A2_CGC IN ({})) SELECT A2_CGC, D1_CC FROM CTE WHERE rn = 1".format(vlist)
    cursor = conn.cursor()  
    cursor.execute(query)
    row = cursor.fetchall()
    result = {}
    for x in row:
        result[x[0]] = x[1]
    return result

def chave_acesso(df):
    vlist = ','.join(["'%s'" % v for v in df[df.columns[0]].values])
    return vlist    

# C T E

def buscarEmitente(df):
    conn = pymssql.connect(server='', user='', password='', database='')      
    vlist = ','.join(["'%s'" % v for v in df[df.columns[1]]])
    query = "SELECT A2_CGC, RTRIM(A2_NOME) FROM SA2010 WHERE A2_CGC IN({})".format(vlist)
    cursor = conn.cursor()  
    cursor.execute(query)
    row = cursor.fetchall()
    result = {}
    for x in row:
        result[x[0]] = x[1]
    return result

def buscarCnpj_Cte(df):
    conn = pymssql.connect(server='', user='', password='', database='')  
    vlist = ','.join(["'%s'" % v for v in df[df.columns[3]]])
    query = "SELECT A2_CGC, RTRIM(A2_NOME) FROM SA2010 WHERE A2_CGC IN({})".format(vlist)
    #print(query)
    cursor = conn.cursor()  
    cursor.execute(query)
    row = cursor.fetchall()
    result = {}
    for x in row:
        result[x[0]] = x[1]
    return result

