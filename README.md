# DocxCast

> **"Your template is the mold. Data is poured in, documents take shape."**

### 🪄 About

**DocxCast** turns a Word document's **content controls** into a typed schema, and casts your data back into a formatted document.

**The Idea:**

A `.docx` template is a natural schema editor. In Word's Developer tab, anyone can insert content controls and describe them:

| Word control property | Role in DocxCast |
|---|---|
| **Title** | field name (e.g. `name`) |
| **Tag** | extraction requirement for the LLM (e.g. `the candidate's full name`) |
| **Repeating section** | array of records |
| **Building block gallery** | `Section` with its own extraction rule |
| **Group control** with a `name?` Title | conditional block (see below) |
| Locked control (`sdtLocked`) | required field |

The full control inventory — every construct's Word-visible pseudo-structure and the OOXML it is made of — lives in **[SYNTAX.md](SYNTAX.md)**.

**Two independent capabilities:**

1. `derive_schema(template)` — read the controls and produce a `Schema`, serializable to JSON Schema for LLM structured extraction.
2. `render(template, data)` — pour data back into the template: clone repeating sections, fill values, unwrap controls, keeping every run's formatting.

### ⚡ Key Features

* **Your template IS the schema** — field names, extraction requirements and repeats all live in the docx, editable by non-programmers.
* **Typed control mapping** — text → `string`, dropdown → `enum`, date / date-time → `format`, checkbox → `boolean`, picture → derived and removed.
* **LLM-friendly contract validation** — `schema.validate(data)` returns structured, machine-readable issues (`path` / `code` / `message` / `expected` / `got`) that can be fed back into an LLM correction loop.
* **Lenient by default, strict on demand** — missing fields keep their template values; `strict=True` raises on the first error.
* **Clean output** — controls are unwrapped by default; `keep_controls=True` keeps them for round-trip editing.
* **Zero magic** — a single runtime dependency (`python-docx`).

### 🚀 Quick Start

```python
from docxcast import derive_schema, render

schema = derive_schema('resume-template.docx')

# hand this to your LLM as the extraction contract
json_schema = schema.to_json_schema()

# validate what the LLM extracted
report = schema.validate(data)
if not report.ok:
    # feed report.issues back into the LLM for correction
    ...

# cast data into the template; controls are unwrapped
result = render('resume-template.docx', data)
result.save('resume-output.docx')
```

### 📏 Validation

`Schema.validate()` never raises; it returns a `ValidationResult`:

```python
result.ok      # False if any error-level issue exists
result.errors  # missing_required / type_mismatch / enum_invalid / format_invalid
result.warnings  # unknown_key
```

`render(..., strict=True)` raises `RenderError` on the first error-level issue.

### 🔀 Conditional Blocks

Wrap a block in a **group control** and mark the guarded control's Title with a trailing `?` (TS-style): a falsy value removes the whole group — label text included — while a truthy value fills normally.

```
Nationality: [nationality?]   truthy → "Nationality: Chinese"   falsy → the line is gone

Work Experiences              both the heading and the repeat
[repeat: work_experiences?]   disappear when the array is empty
```

The `?` is a render directive only — the derived schema (and the LLM contract) see a plain `nationality` field. Marker placement, inline vs block-level drops, positional `has_prev?` / `has_next?` separators and every misuse error live in **[SYNTAX.md](SYNTAX.md)**.

### ⚠️ Limitations

* `picture` controls are recognized in the schema but removed on render (no image insertion yet).
* Date values are written verbatim; the template's display mask is not enforced.
* Unnamed controls are skipped in the schema — every field needs a Title to be addressable.
* No whitespace control: text outside a dropped inline conditional span survives — put separators inside the span.
