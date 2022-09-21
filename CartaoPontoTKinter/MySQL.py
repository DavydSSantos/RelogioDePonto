import mysql.connector
from Config import *

mydb = mysql.connector.connect(

	host     = IPSQL,
	user     = UsuarioBanco,
	password = SenhaBanco,
    database = NomeBanco
)
