# Versioning

Both packages follow semver, and are released together at the same version:
`@xlscalc/xlsx-preview` depends on `@xlscalc/formula-engine` and the two are
tested as a pair. While the major is `0`, a **minor** bump may break API.

That much is ordinary. The part that is not ordinary is below.

## Capability growth is a breaking change to output

Most libraries can add a feature without changing what existing callers see.
This one cannot, and pretending otherwise would be the most damaging kind of
quiet release.

A cell whose formula this engine does not support renders `⚠`, and refusal
propagates: everything downstream of it renders `⚠` too. So the day `OFFSET` is
implemented, a workbook that used to show a screen of warnings shows a screen of
numbers. Nobody's code changed. The library did not break. But every screenshot,
every cached render, every snapshot test and every "this file is not supported"
routing decision a host made is now wrong.

In the real corpus, 33 refusing cells darken 64,809 — a little under half a
workbook. One function is not a footnote.

So:

- **Adding or removing a function is at least a minor bump**, never a patch, and
  it is named in the changelog under a *Capability* heading with the functions
  listed. If you pin, pin to a patch range.
- **`CAPABILITY.md` is generated from the registry**, not written by hand, and it
  ships inside `@xlscalc/formula-engine`. Diffing it between two installed
  versions is the supported way to see exactly what changed about what renders,
  which means it has to be in the tarball and not only in the repository.
  `implementedFunctions()` and `refusedFunctions()` return the same lists at
  runtime.
- **A fix that changes a computed value is also at least a minor bump**, for the
  same reason and with the same heading. Agreeing with Excel more closely than
  we did last week is still a different number on someone's screen.

## What is *not* covered by semver

- **The real-data corpus is ten workbooks, and it is private.** Wide ones, from
  five different writers, but not a random sample — and they are real businesses'
  confidential files, so a reader cannot rerun that gate. A release being green
  means it agrees with every oracle we have, not that it agrees with Excel
  everywhere.
- **Declared divergences.** The real corpus's rule file lists the places we
  knowingly differ from a file's own cached values, each with an exact expected
  count. They are documented rather than fixed, and the count is gated in both
  directions — a rule that starts explaining more cells, or fewer, fails the
  build. That file is private along with the corpus itself; the README says why.
- **`stylesError` output.** The fidelity of the *styling* parse depends on
  ExcelJS. A workbook it cannot read still renders its values, in default fonts.

## Releasing

Order matters, because the second package depends on the first:

```bash
npm run verify                              # typecheck, tests, build, and the published-package smoke
npm publish -w @xlscalc/formula-engine      # prepack rebuilds dist/ and copies CAPABILITY.md in
npm publish -w @xlscalc/xlsx-preview        # its dependency range must already resolve on the registry
```

Both are scoped, so the `@xlscalc` npm organisation has to exist and the
publishing account has to be a member of it; `publishConfig.access` is already
`public`, which scoped packages need or npm assumes private. `prepack` runs the
build, so a stale `dist/` cannot ship — but nothing rebuilds `CAPABILITY.md`, so
run `npm test` first and let it fail if the registry has moved.

`apps/demo` is `private: true` and is not published.

## Deprecations

Anything deprecated keeps working for one minor release with a runtime warning,
and is removed in the next. Deprecated exports are listed in the changelog when
they are deprecated *and* when they are removed.
