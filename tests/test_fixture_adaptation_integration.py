"""Functional regression tests for fixture adaptation using an isolated in-memory SQLite DB."""
import os
import unittest
from datetime import date, timedelta

# Must be set before importing the application; never use production DATABASE_URL.
os.environ["DATABASE_URL"] = "sqlite:///:memory:"
os.environ.setdefault("SECRET_KEY", "fixture-tests-only")

import app as sistema


class FixtureAdaptationIntegrationTests(unittest.TestCase):
    def setUp(self):
        self.app_context = sistema.app.app_context()
        self.app_context.push()
        sistema.db.drop_all()
        sistema.db.create_all()

        self.clubes = [
            sistema.Club(nombre=f"Club prueba {i}")
            for i in range(1, 6)
        ]
        sistema.db.session.add_all(self.clubes)
        sistema.db.session.flush()

        self.campeonato = sistema.Campeonato(
            nombre="Campeonato prueba integración",
            temporada="2026",
            serie="Senior",
            fecha_inicio=date(2026, 5, 16),
            estado="Activo",
        )
        sistema.db.session.add(self.campeonato)
        sistema.db.session.flush()

        for club in self.clubes:
            sistema.db.session.add(
                sistema.CampeonatoClub(
                    campeonato_id=self.campeonato.id,
                    club_id=club.id,
                )
            )
        sistema.db.session.flush()

        calendario = sistema.generar_calendario_todos_contra_todos(
            [club.id for club in self.clubes]
        )
        fecha_base = date(2026, 5, 16)
        for jornada, cruces in enumerate(calendario, start=1):
            for local_id, visitante_id in cruces:
                finalizado = jornada == 1
                sistema.db.session.add(
                    sistema.Partido(
                        campeonato_id=self.campeonato.id,
                        jornada=jornada,
                        fecha=fecha_base + timedelta(days=(jornada - 1) * 7),
                        hora="17:00",
                        cancha=f"Cancha {jornada}",
                        local_club_id=local_id,
                        visitante_club_id=visitante_id,
                        goles_local=2 if finalizado else None,
                        goles_visitante=1 if finalizado else None,
                        estado="Finalizado" if finalizado else "Programado",
                        fixture_bloqueado=False,
                    )
                )
        sistema.db.session.commit()

    def tearDown(self):
        sistema.db.session.remove()
        sistema.db.drop_all()
        self.app_context.pop()

    def _snapshot(self):
        return [
            (
                p.id, p.jornada, p.fecha, p.hora, p.cancha,
                p.local_club_id, p.visitante_club_id,
                p.goles_local, p.goles_visitante, p.estado,
                p.fixture_bloqueado,
            )
            for p in sistema.Partido.query.filter_by(
                campeonato_id=self.campeonato.id
            ).order_by(sistema.Partido.jornada, sistema.Partido.id).all()
        ]

    def test_adaptation_preserves_played_round_and_unique_matchups(self):
        before_played = [
            row for row in self._snapshot() if row[1] == 1
        ]
        trigger = type("JornadaDisparador", (), {"jornada": 2})()

        sistema.adaptar_fixture_pendiente(
            self.campeonato,
            trigger,
            proteger_jornada_modificada=False,
        )
        sistema.db.session.commit()

        after = self._snapshot()
        after_played = [row for row in after if row[1] == 1]
        self.assertEqual(after_played, before_played)

        pairs = [
            frozenset((row[5], row[6]))
            for row in after
        ]
        self.assertEqual(len(pairs), 10, "A 5-team single round-robin must have 10 matches")
        self.assertEqual(len(set(pairs)), len(pairs), "No matchup may be repeated")

        for jornada in range(1, 5 + 1):
            matches = [row for row in after if row[1] == jornada]
            participants = [club_id for row in matches for club_id in (row[5], row[6])]
            self.assertEqual(len(participants), len(set(participants)),
                             f"A club appears more than once in round {jornada}")
            self.assertEqual(len(set(self.clubes_ids) - set(participants)), 1,
                             f"Round {jornada} must have exactly one bye")

    @property
    def clubes_ids(self):
        return [club.id for club in self.clubes]

    def test_future_locked_match_aborts_without_any_change(self):
        locked = sistema.Partido.query.filter_by(
            campeonato_id=self.campeonato.id,
            jornada=3,
        ).first()
        locked.fixture_bloqueado = True
        sistema.db.session.commit()
        before = self._snapshot()

        trigger = type("JornadaDisparador", (), {"jornada": 2})()
        with self.assertRaisesRegex(ValueError, "protegidos manualmente"):
            sistema.adaptar_fixture_pendiente(
                self.campeonato,
                trigger,
                proteger_jornada_modificada=False,
            )
        sistema.db.session.rollback()

        self.assertEqual(self._snapshot(), before,
                         "A refused adaptation must not alter any fixture row")


if __name__ == "__main__":
    unittest.main()
