# Image test results — NI Report Generation Toolkit

Italian version: [IMAGE_TEST_RESULTS-IT.md](IMAGE_TEST_RESULTS-IT.md).

This document records the Open XML behavior observed while characterizing the NI Excel image VIs. The Italian version retains the complete original test notes; this English version preserves the tested inputs, measurements, conclusions, and resulting backend contract.

## References

- Baseline: `../LV/Pre-Test/IMAGE_POSITION_TEST_BASELINE.xlsx`
- Image: `../LV/Pre-Test/prova.png`
- RGB PNG, 320 × 200 pixels, declared 300 DPI, 724 bytes

## Characterized interfaces

`Append Image to Report (string wrap).vi` accepts the MS Office position bundle, an NI alignment enum, the report reference, image path, caption/alternative string, and error cluster. It has no width or height inputs; sizing belongs to `Excel Format Image.vi`.

`Excel Format Image.vi` accepts measurement system (`US` or `metric`), scale factor, report reference, picture index (`-1` means the last image), height, width, color type, and error cluster.

## FMT-01 — scale factor 0.5

Observed in `FMT-01.xlsx` with metric units, picture index `-1`, dimensions `-1`, and automatic color:

- original size: approximately 320.008 × 200.005 px;
- result: approximately 160.004 × 100.003 px;
- exact 0.5 ratio on both axes with aspect ratio preserved;
- embedded PNG unchanged, including size and SHA-256;
- only the DrawingML anchor and transform changed;
- anchor remained `twoCellAnchor editAs="oneCell"` at `A1`.

Contract: a positive scale factor multiplies both current dimensions, `1` leaves the size unchanged, the anchor origin is preserved, and the media part is not transcoded. Zero or negative scale factors should be rejected unless a concrete application requirement says otherwise.

## FMT-02 / FMT-03 / FMT-03m — explicit dimensions

| File | System | Width | Height | Observed physical result |
|---|---|---:|---:|---|
| `FMT-02.xlsx` | US | 160 | -1 | 160 in × 0 |
| `FMT-03.xlsx` | US | -1 | 100 | 0 × 100 in |
| `FMT-03m.xlsx` | metric | -1 | 100 | 0 × approximately 100.0125 cm |

The NI implementation treats US values as inches and metric values as centimetres. When only one dimension is positive, its resize branch incorrectly collapses the other `-1` dimension to zero. When both dimensions are `-1`, geometry remains unchanged.

The new backend intentionally implements the useful documented semantics: a negative value preserves that dimension, a nonnegative value is converted from the selected unit, and explicit zero is rejected unless required later. This is a deliberate correction of the observed NI partial-resize bug.

## FMT-04 — explicit metric width and height

For inputs `width=160`, `height=100`, the resulting geometry was:

- width 57,607,200 EMU = 4,536 pt = 160.02 cm;
- height 36,004,500 EMU = 2,835 pt = 100.0125 cm;
- ratio 1.6, origin `A1`, unchanged embedded PNG.

The observed NI conversion quantizes through integer Office points:

```text
US:     points = round(inches * 72)
metric: points = round(centimetres / 2.54 * 72)
EMU:    points * 12,700
```

The backend reproduces this quantization for geometric compatibility rather than using a more precise direct centimetre-to-EMU conversion.

## FMT-GS — grayscale

`FMT-GS.xlsx` leaves the media bytes, hash, geometry, and relationship unchanged. It adds one `<a:grayscl/>` child to the selected picture's `a:blip`. Grayscale is therefore a nondestructive rendering effect belonging to the picture object, not to the shared media part.

Automatic color removes transformations applied by this function. `msoPictureMixed` is an Office mixed-state value rather than a concrete transformation and should not be accepted as a command without a specific requirement. Black-and-white and watermark remain outside the implemented contract.

## FMT-PI — picture-index selection

`FMT-PI.xlsx` contains the same 320 × 311 PNG at `A1` and `H7`. Formatting index `0` selects the first object; index `1` selects the second. Only index `1` receives grayscale, even though both objects share a single `ImagePart`.

Picture indexes are zero-based and follow drawing insertion order. `-1` selects the last image. Values below `-1` or greater than or equal to the picture count must produce a deterministic error.

## IMG-01 — default insertion

Observed geometry:

- modified worksheet: `Sheet1`;
- `xdr:twoCellAnchor editAs="oneCell"`;
- start marker `A1`, zero offsets;
- end marker column 5, row 10, with 76 EMU X and 95,298 EMU Y offsets;
- transform `cx=3,048,076`, `cy=1,905,048` EMU;
- approximately 320.008 × 200.005 px with aspect ratio preserved;
- `a14:useLocalDpi val="0"`.

The toolkit uses the original pixel size at 96 logical DPI and ignores the PNG's declared 300 DPI. The embedded `xl/media/image1.png` is byte-identical to the source: 724 bytes, SHA-256 `4B2C54ECC20EB1D0158A6B62C953854CBD54D01B0A138F44C7BB35327BDB87EB`.

NI creates both an embedded relationship and an external relationship to the absolute source path. The external link is unnecessary for display, introduces an environment dependency, and leaks the generating machine's path. XCelReportEngine intentionally creates only the embedded relationship.

Required reproduction contract:

1. embed original bytes without transcoding;
2. create `TwoCellAnchor` with `EditAs=OneCell`;
3. use zero origin offsets;
4. convert pixels at approximately 9,525 EMU/px;
5. compute the end marker across actual row and column geometry;
6. set `useLocalDpi=false`;
7. omit the external source relationship.

## IMG-03 — position C5

Inputs `(row=4, column=2)` produce a start marker at zero-based column 2, row 4: Excel cell `C5`. Size and media remain unchanged. This confirms that NI numeric coordinates are ordered `(row, column)` and are both zero-based:

- `(0, 0)` → `A1`;
- `(2, 4)` → `E3`;
- `(4, 2)` → `C5`.

These values map directly to DrawingML marker indexes without an additional `+1/-1` conversion.

## IMG-04 — RIGHT alignment

Changing alignment to `RIGHT` while keeping the `C5` position and image produced no functional difference: anchor, offsets, size, transform, media, and relationships were identical. Only Excel's random `a16:creationId` changed.

The Excel backend may therefore accept the alignment parameter for connector-pane compatibility and ignore it.

## IMG-05 — caption/alternative string

Providing a caption caused no drawing, worksheet, shared-string, media, relationship, `descr`, or `title` change. The NI Excel backend ignores this input. XCelReportEngine follows that behavior; a future accessibility enhancement could populate `xdr:cNvPr/@descr`, but that would not be a 1:1 reproduction.

## IMG-06 — A1 cell string

With numeric coordinates `(0, 0)` and `cellAddress="H7"`, the result starts at DrawingML column 7, row 6 (`H7`) with zero offsets. The cell string therefore takes precedence over numeric coordinates.

Contract: validate and use a nonempty A1 string; otherwise use the zero-based numeric pair. An invalid nonempty string must produce a deterministic error rather than silently falling back to numeric coordinates.

## IMG-07 — repeated insertion

Two insertions of the same PNG at `H7` and `A1` create one shared media part and two distinct picture anchors. Both objects reuse the same embedded relationship and preserve insertion order. Object identifiers and names are unique but need not reproduce Excel's cosmetic numbering gaps.

Contract:

- deduplicate identical media bytes within a worksheet;
- create a new anchor/picture for every append call;
- generate unique IDs considering pre-existing drawing objects;
- append anchors in call order so insertion order also defines z-order;
- do not create an external source relationship.
