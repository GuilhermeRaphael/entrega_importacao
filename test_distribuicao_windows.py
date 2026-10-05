import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
import core
import test_gestores

class DistribuicaoWindowsTests(unittest.TestCase):
    def test_nome_matriz_preserva_maiusculas_e_acentos(self):
        with tempfile.TemporaryDirectory() as pasta:
            caminho=Path(pasta)/"MATRIZ - CELULAR INRI ÁREA.xlsx"
            test_gestores.ImportacaoTests().matriz(caminho,14)
            nome_original=caminho.name
            cfg={"matrizes":[{"exibicao":"CELULAR INRI","caminho":str(caminho)}]}
            core.consolidar_remessa(cfg,[core.ItemFila("CELULAR INRI","A","V1",1)])
            self.assertEqual([p.name for p in Path(pasta).iterdir()],[nome_original])
            self.assertTrue(caminho.exists())

    def test_configuracao_fica_ao_lado_do_executavel(self):
        executavel=os.path.join(os.path.abspath("portable"),"ImportadorGoalfy.exe")
        with patch.object(sys,"frozen",True,create=True), patch.object(sys,"executable",executavel):
            self.assertEqual(core.diretorio_aplicacao(),os.path.dirname(executavel))

if __name__=="__main__":
    unittest.main()
