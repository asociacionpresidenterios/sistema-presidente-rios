import ast
import unittest
from pathlib import Path


APP_PATH = Path(__file__).resolve().parents[1] / "app.py"
SOURCE = APP_PATH.read_text(encoding="utf-8")
TREE = ast.parse(SOURCE)


def function_node(name):
    for node in TREE.body:
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)) and node.name == name:
            return node
    raise AssertionError(f"Function not found: {name}")


def source_segment(node):
    return ast.get_source_segment(SOURCE, node)


class FixtureSafetyRegressionTests(unittest.TestCase):
    def test_fixture_generator_refuses_existing_matches(self):
        node = function_node("generar_fixture_campeonato")
        code = source_segment(node)
        guard = code.find("partidos_existentes")
        destructive = code.find(".delete(")
        self.assertGreaterEqual(guard, 0, "Missing guard for existing fixture")
        self.assertIn("if partidos_existentes:", code)
        self.assertIn("No se modificó ningún encuentro", code)
        self.assertTrue(destructive == -1 or guard < destructive)

    def test_adaptation_refuses_future_locked_matches(self):
        node = function_node("adaptar_fixture_pendiente")
        code = source_segment(node)
        guard = code.find("partidos_bloqueados_futuros")
        destructive = code.find("db.session.delete(p)")
        self.assertGreaterEqual(guard, 0, "Missing guard for locked future matches")
        self.assertIn("fixture_bloqueado", code)
        self.assertIn("_partido_finalizado(p)", code)
        self.assertIn("raise ValueError(", code[guard:destructive if destructive >= 0 else None])
        self.assertTrue(destructive == -1 or guard < destructive)

    def test_fixture_deletion_refuses_finalized_or_locked_matches(self):
        node = function_node("eliminar_fixture_campeonato")
        code = source_segment(node)
        guard = code.find("partidos_protegidos")
        destructive = code.find(".delete(")
        self.assertGreaterEqual(guard, 0, "Missing protected-match check")
        self.assertIn("_partido_finalizado(p)", code)
        self.assertIn("fixture_bloqueado", code)
        self.assertIn("raise ValueError(", code[guard:destructive if destructive >= 0 else None])
        self.assertTrue(destructive == -1 or guard < destructive)


if __name__ == "__main__":
    unittest.main()
