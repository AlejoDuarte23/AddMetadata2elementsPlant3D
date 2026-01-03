import os
import subprocess
from datetime import datetime
from pathlib import Path


def apply_metadata_rev(
    project_xml: str | os.PathLike,
    json_in: str | os.PathLike,
    plugin_dll: str | os.PathLike,
    *,
    acad_exe: str | os.PathLike = r"C:\Program Files\Autodesk\AutoCAD 2026\acad.exe",
    workdir: str | os.PathLike = r"C:\PlantAutomationRun",
    save_changes: bool = True,
) -> int:
    """Run AutoCAD Plant 3D and apply metadata updates from JSON using Project.xml."""

    acad = Path(acad_exe)
    project_xml = Path(project_xml)
    json_in = Path(json_in)
    plugin_dll = Path(plugin_dll)
    workdir = Path(workdir)

    if not acad.exists():
        raise FileNotFoundError(str(acad))
    if not project_xml.exists():
        raise FileNotFoundError(str(project_xml))
    if not json_in.exists():
        raise FileNotFoundError(str(json_in))
    if not plugin_dll.exists():
        raise FileNotFoundError(str(plugin_dll))

    workdir.mkdir(parents=True, exist_ok=True)

    scr_path = workdir / "run_rev2.scr"
    scr_text = "\n".join(
        [
            "_.NETLOAD",
            f"\"{plugin_dll}\"",
            "_.P3D_APPLY_JSON_METADATA_XML",
            "_.QUIT",
            "Y",
        ]
    ) + "\n"
    scr_path.write_text(scr_text, encoding="utf-8")

    env = os.environ.copy()
    env["PLANT_PROJECT_XML"] = str(project_xml)
    env["PLANT_JSON_IN"] = str(json_in)
    env["PLANT_SAVE_CHANGES"] = "1" if save_changes else "0"
    log_path = workdir / f"plant_run_{datetime.now():%Y%m%d_%H%M%S}.jsonl"
    env["PLANT_LOG_PATH"] = str(log_path)

    cmd = [
        str(acad),
        "/product",
        "PLNT3D",
        "/b",
        str(scr_path),
    ]

    completed = subprocess.run(cmd, env=env, cwd=str(workdir))
    return completed.returncode


def default_plugin_dll() -> Path:
    repo_root = Path.cwd().resolve()
    return (
        repo_root
        / "addin"
        / "bin"
        / "Release"
        / "net8.0-windows"
        / "MetadataApplier.dll"
    )


def default_json() -> Path:
    repo_root = Path.cwd().resolve()
    return repo_root / "sample" / "metadata.json"


def main() -> int:
    project_xml = os.environ.get(
        "PLANT_PROJECT_XML",
        r"C:\Users\aleja\Downloads\NIRAS P3D Demo Project\P3D-FBUK Example Project\Project.xml",
    )
    json_in = os.environ.get("PLANT_JSON_IN", str(default_json()))
    plugin_dll = os.environ.get("PLANT_PLUGIN_DLL", str(default_plugin_dll()))

    if not project_xml:
        raise SystemExit("PLANT_PROJECT_XML is required.")

    return apply_metadata_rev(project_xml, json_in, plugin_dll)


if __name__ == "__main__":
    raise SystemExit(main())
