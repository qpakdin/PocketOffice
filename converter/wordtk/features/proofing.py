# -*- coding: utf-8 -*-
"""Proofing group: hide spelling/grammar squiggles for this document, mark text as not-to-be-checked."""
from ..core.ooxml import ensure_rpr, qn, set_onoff, is_on
from ..core.scope import iter_paragraphs
from ..core.textops import paragraph_runs
from .base import F, Feature


class ProofingFeature(Feature):
    key = "proofing"
    label = "Proofing"
    description = "Hide spelling / grammar errors in this document only; mark text as 'do not check'."
    fields = [
        F("hide_spelling_errors", "bool", "Hide spelling errors in this document only"),
        F("hide_grammar_errors", "bool", "Hide grammar errors in this document only"),
        F("no_proof_text", "bool", "Do not check spelling or grammar (mark selected text)",
          hint="writes w:noProof on every run in scope; false removes it"),
    ]

    def apply(self, ctx, s):
        st = ctx.pkg.xml("word/settings.xml", edit=True) if ctx.pkg.exists("word/settings.xml") else None
        for key, tag in (("hide_spelling_errors", "w:hideSpellingErrors"),
                         ("hide_grammar_errors", "w:hideGrammaticalErrors")):
            if key in s:
                if st is None:
                    ctx.warn("proofing: no settings.xml; %s skipped" % key)
                else:
                    set_onoff(st, tag, True if s[key] else None)
                    ctx.count("proofing." + key)
        if "no_proof_text" in s:
            scope, n = self.scope_for(ctx, s), 0
            for _n, _r, p in iter_paragraphs(ctx, scope, edit=True):
                for r in paragraph_runs(p):
                    set_onoff(ensure_rpr(r), "w:noProof", True if s["no_proof_text"] else None)
                    n += 1
            ctx.count("proofing.runs", n)

    def read(self, pkg, include_text=False):
        st = pkg.try_xml("word/settings.xml")
        return {"hide_spelling_errors": st is not None and st.find(qn("w:hideSpellingErrors")) is not None,
                "hide_grammar_errors": st is not None and st.find(qn("w:hideGrammaticalErrors")) is not None,
                "no_proof_runs": sum(1 for name, _k in pkg.story_parts()
                                     for _ in pkg.xml(name).iter(qn("w:noProof")))}

    def snapshot(self, pkg):
        r = self.read(pkg)
        return {k: r[k] for k in ("hide_spelling_errors", "hide_grammar_errors") if r[k]}
