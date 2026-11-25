# -*- coding: utf-8 -*-
"""Batch processing of many DOCX files with progress/cancel support."""
import time
from pathlib import Path

from ..core.package import DocxError
from ..features import ProfileError
from .profile import apply_profile_to_file


def collect_files(paths, recursive=True):
    """Expand files/folders into [(file, root_folder_or_None)]; skips ~$ lock files and duplicates."""
    out, seen = [], set()
    for p in map(Path, paths):
        if p.is_dir():
            it = p.rglob("*.docx") if recursive else p.glob("*.docx")
            cands = [(f, p) for f in sorted(it)]
        elif p.suffix.lower() == ".docx" and p.is_file():
            cands = [(p, None)]
        else:
            cands = []
        for f, root in cands:
            if f.name.startswith("~$"):
                continue
            k = str(f.resolve())
            if k not in seen:
                seen.add(k)
                out.append((f, root))
    return out


def output_path(src, root, out_dir, preserve=True, taken=None):
    """Output file keeps exactly the input file's name (same stem and extension).

    A file left in the output folder by an earlier run is replaced. Only if two inputs of the
    *same run* would land on the same path (e.g. flat mode, same name in different folders) the
    later one gets a numeric counter so it never clobbers a file written moments before.
    """
    src = Path(src)
    rel = src.relative_to(root) if (root is not None and preserve) else Path(src.name)
    dst = Path(out_dir) / rel.parent / rel.name
    if taken is not None:
        i = 2
        while str(dst).lower() in taken:
            dst = dst.with_name("%s_%d%s" % (rel.stem, i, rel.suffix))
            i += 1
        taken.add(str(dst).lower())
    return dst


def run_batch(files, profile, out_dir, preserve=True, dry_run=False,
              on_progress=None, should_stop=None):
    """files: [(path, root)]. Returns list of result dicts."""
    results = []
    taken = set()
    for n, (src, root) in enumerate(files, 1):
        if should_stop and should_stop():
            break
        t0 = time.time()
        res = {"source": str(src), "status": "ok", "output": None, "changes": {}, "warnings": []}
        try:
            dst = output_path(src, root, out_dir, preserve, taken)
            if dst.resolve() == Path(src).resolve():
                raise DocxError("output equals input")
            r = apply_profile_to_file(src, dst, profile, dry_run)
            res.update(r)
            res["output"] = None if dry_run else str(dst)
        except (DocxError, ProfileError) as e:
            res.update(status="error", error=str(e))
        except Exception as e:  # keep the batch going
            res.update(status="error", error="%s: %s" % (type(e).__name__, e))
        res["seconds"] = round(time.time() - t0, 2)
        results.append(res)
        if on_progress:
            on_progress(n, len(files), res)
    return results
