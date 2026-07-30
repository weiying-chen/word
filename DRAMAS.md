# Drama DOCX generation

`generate_drama.py` renders a Schema Version 2.0 bilingual drama JSON as an
English shooting-script DOCX. It uses the Chinese-source record order and a
manager-produced English DOCX for screenplay structure and page layout. Font,
text size, and paragraph spacing come from this repository's shared document
style system. It never translates text or falls back to Chinese.

## Preview

Use preview mode while translation is incomplete:

```bash
gen-dramas \
  --input "/home/weiying/text/dramas/sy/working/靜曦拍攝本V.4.2.translation.json" \
  --reference "/home/weiying/text/dramas/sy/refs/他的靈魂與他的書店第8集_劇本_英文.docx" \
  --verify-source "/home/weiying/text/dramas/sy/refs/靜曦拍攝本V.4.2.pdf" \
  --preview
```

The preview is visibly marked incomplete, emits only translated records in
global order, and reports the number of required English fields still missing.
Its default name is `靜曦拍攝本V.4.2_英文.preview.docx`.

## Final document

Omit `--preview` for final generation:

```bash
gen-dramas \
  --input "/path/to/completed.translation.json" \
  --reference "/path/to/manager-reference.docx" \
  --verify-source "/path/to/source.pdf" \
  --output "/optional/custom/final.docx"
```

Final generation refuses to write a document when the schema, IDs, global
order, scene numbering, English fields, source hash, or review flags fail
validation. Without `--output`, the final document is written beside the JSON
with `_英文.docx` appended to its base name.

`merged_visual_continuation` is informational extraction provenance and does
not block final output. Ambiguity flags such as
`possible_dialogue_continuation`, plus unknown review flags, remain blocking.

Each source record is mapped to one hidden DOCX bookmark. The generator checks
the saved package and mapping before reporting success.

## Commands

- `gen-dramas`: generate a drama DOCX with this repository.
- `clean-dramas`: run the older TXT cleanup tool that removes blank lines.

The Python entry point can also be called directly with
`python3 generate_drama.py`.
