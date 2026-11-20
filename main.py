from fastapi import FastAPI
from pydantic import BaseModel
from randomGaus import cpk_estudio
from pydantic import BaseModel
from typing import Optional, Literal


class CPKResponse(BaseModel):
    status: Literal["success", "error"]
    message: str
    project: str
    analysis: str
    folder_url: Optional[str] = None
    error: Optional[str] = None


class cpk(BaseModel):
    proyecto: str
    equipo: str
    numero: int
    hoja: int
    upl: float
    lwl: float

app = FastAPI()

@app.post("/CPK", response_model=CPKResponse)
def crear_CPK(CPK: cpk):
    try:
        resultado = cpk_estudio(CPK.upl, CPK.lwl)
        if resultado.startswith("Estudio erroneo"):
            status = "error"
            error = resultado
        else:
            status = "success"
            error = ""
        return CPKResponse(
            status = status,
            message=resultado,
            project=CPK.proyecto,
            analysis="CPK",
            folder_url=None,
            error=error
        )

    except Exception as e:
        return CPKResponse(
            message="No fue posible completar el estudio",
            project=CPK.proyecto,
            analysis="CPK",
            folder_url=None,
            error=str(e)
        )

    # return {lwl, uwl, f"Proyecto: {CPK.proyecto}, equipo: {CPK.equipo}{CPK.numero}, hoja:{CPK.hoja}"}

#print(crear_CPK)