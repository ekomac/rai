from dataclasses import dataclass
import unidecode


@dataclass
class Pregunta:

    TYPE_BOL = "BOL"
    TYPE_VEM = "VEM"
    TYPE_VEX = "VEX"
    TYPE_VEC = "VEC"
    TYPE_MULT = "MULT"
    TYPE_SING = "SING"
    TYPE_UNDEFINED = "undefined"

    id: str
    pregunta: str
    respuestas: str
    categoria: str
    cant_respuestas: int = 0
    tipo: str = TYPE_UNDEFINED

    @property
    def respuestas_as_list(self):
        if self.respuestas is None:
            return []
        return list(map(
            lambda x: int(x) if x.isdigit() else x,
            self.respuestas.split("; ")
        ))

    def __post_init__(self):
        self.cant_respuestas = len(
            self.respuestas_as_list) if self.respuestas else 0
        self.__set_tipo()
        self.__update_respuestas_if_none()

    def __set_tipo(self):
        if self.cant_respuestas == 2:
            if all(x in self.respuestas_as_list for x in ["Si", "No"]):
                self.tipo = self.TYPE_BOL
                return
        if self.cant_respuestas == 5:
            if all(x in self.respuestas_as_list for x in list(range(1, 6))):
                self.tipo = self.TYPE_VEM
                return
        if self.cant_respuestas > 5:
            if all(x in self.respuestas_as_list for x in list(range(1, 6))):
                self.tipo = self.TYPE_VEX
                return
        if self.cant_respuestas >= 3:
            count = 0
            vecs = ['mas de', 'mas que' 'entre',
                    'hace', 'menos de', 'menos que',
                    'nunca', 'siempre', ]
            for rta in self.respuestas_as_list:
                if isinstance(rta, str):
                    resp = unidecode.unidecode(rta.lower())
                    if any(x in resp for x in vecs):
                        count += 1
                    if count > 1:
                        self.tipo = self.TYPE_VEC
                        return

    def __update_respuestas_if_none(self):
        if self.respuestas is None:
            self.respuestas = ""
