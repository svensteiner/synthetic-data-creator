"""
run_all.py - Alle Tests starten und Ergebnis uebersichtlich anzeigen.
"""

import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(Path(__file__).parent))

if __name__ == "__main__":
    loader = unittest.TestLoader()
    suite  = loader.discover(start_dir=str(Path(__file__).parent),
                              pattern="test_*.py")

    runner = unittest.TextTestRunner(verbosity=2, stream=sys.stdout)
    result = runner.run(suite)

    print()
    print("=" * 60)
    if result.wasSuccessful():
        print(f"  ALLE {result.testsRun} TESTS BESTANDEN")
    else:
        print(f"  {result.testsRun} Tests, "
              f"{len(result.failures)} Fehler, "
              f"{len(result.errors)} Exceptions")
    print("=" * 60)

    sys.exit(0 if result.wasSuccessful() else 1)
