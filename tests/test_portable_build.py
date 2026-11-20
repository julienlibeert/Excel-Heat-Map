import json
from pathlib import Path
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import MagicMock, patch

from excel_heat_map.utils.portable_build import collect_runtime


class RuntimeCollectionTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name)
        self.base = self.root / "python"
        self.dlls = self.base / "DLLs"
        self.runtime = self.base / "Library" / "bin"
        self.system = self.root / "Windows" / "System32"
        self.cache = self.root / "cache"
        for directory in (self.dlls, self.runtime, self.system):
            directory.mkdir(parents=True)
        self.extension = self.dlls / "_tkinter.pyd"
        self.tk = self.runtime / "tk86t.dll"
        self.zlib = self.runtime / "zlib.dll"
        for path in (self.extension, self.tk, self.zlib, self.system / "kernel32.dll"):
            path.write_bytes(path.name.encode())
        self.imports = {
            self.extension: (
                ["tk86t.dll", "kernel32.dll", "api-ms-win-core-test.dll"],
                [],
            ),
            self.tk: ([], ["zlib.dll"]),
            self.zlib: ([], []),
        }

    def collect(self):
        def image(path, fast_load):
            imports, delayed = self.imports[Path(path)]
            context = MagicMock()
            context.__enter__.return_value = SimpleNamespace(
                DIRECTORY_ENTRY_IMPORT=[
                    SimpleNamespace(dll=name.encode()) for name in imports
                ],
                DIRECTORY_ENTRY_DELAY_IMPORT=[
                    SimpleNamespace(dll=name.encode()) for name in delayed
                ],
                parse_data_directories=lambda **_kwargs: None,
            )
            return context

        pefile = SimpleNamespace(
            PE=image,
            DIRECTORY_ENTRY={
                "IMAGE_DIRECTORY_ENTRY_IMPORT": 1,
                "IMAGE_DIRECTORY_ENTRY_DELAY_IMPORT": 13,
            },
        )
        with (
            patch("sys.base_prefix", str(self.base)),
            patch.dict("os.environ", {"SystemRoot": str(self.system.parent)}),
            patch.dict("sys.modules", {"pefile": pefile}),
        ):
            return collect_runtime(self.cache)

    def test_collects_recursive_and_delayed_dependencies_but_not_windows_dlls(self):
        names = self.collect()
        self.assertCountEqual(names, ["tk86t.dll", "zlib.dll"])
        self.assertEqual((self.cache / "zlib.dll").read_bytes(), self.zlib.read_bytes())
        self.assertFalse((self.cache / "kernel32.dll").exists())
        manifest = json.loads((self.cache / "manifest.json").read_text())
        self.assertCountEqual(manifest["files"], names)

    def test_refreshes_changed_dlls_and_excludes_obsolete_cached_files(self):
        self.collect()
        (self.cache / "obsolete.dll").write_bytes(b"old dependency")
        self.zlib.write_bytes(b"updated runtime")
        names = self.collect()
        self.assertEqual((self.cache / "zlib.dll").read_bytes(), b"updated runtime")
        self.assertNotIn("obsolete.dll", names)

    def test_unresolved_non_system_dependency_fails_explicitly(self):
        self.zlib.unlink()
        with self.assertRaisesRegex(FileNotFoundError, r"zlib\.dll.*tk86t\.dll"):
            self.collect()


if __name__ == "__main__":
    unittest.main()
