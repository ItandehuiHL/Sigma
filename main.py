from fastapi import FastAPI
from pydantic import BaseModel
from randomGaus import cpk_estudio
import win32com.client

class cpk(BaseModel):
    proyecto: str
    equipo: str
    numero: int
    hoja: int
    upl: float
    lwl: float

app = FastAPI()

@app.post("/CPK")
def crear_CPK(CPK: cpk):
    return cpk_estudio(CPK.upl, CPK.lwl)
    # return {lwl, uwl, f"Proyecto: {CPK.proyecto}, equipo: {CPK.equipo}{CPK.numero}, hoja:{CPK.hoja}"}

#print(crear_CPK)