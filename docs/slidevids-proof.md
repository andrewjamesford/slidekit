# SlideVids fork prototype

Version `0.2.0-sv.1`, MIT. Upstream v0.2.0 commit `36131748efbd0ca952d0630ddb3a0557b848bcff`.

Schema 2 adds `slide.notes: string`, maximum 100,000 UTF-16 code units. `validateDocument` upgrades schema 1 into an owned copy with empty Notes, preserving slide/object IDs and geometry. No canonical timing field exists yet. Upstream schema 1 examples still load, but exported schema 2 cannot be opened by the unchanged upstream release.

The public editor `addSlide` and `updateSlide` accept Notes. `moveElement(id, targetIndex)` moves unlocked current-slide elements to a zero-based back-to-front paint index as one history transaction. Arrange displays frontmost first and supports dragging and forward/backward buttons. View mode disables Notes and layer mutations. Notes is the only editable narration field; no provider is called. Localization, IME and further accessibility review remain for T06.

PPTX import reads notesSlide relationships, body paragraphs, run text, soft breaks and Unicode. Plain-text export uses PptxGenJS Notes support. Import findings state rich Notes/layout losses. Generated rich-text PPTX paragraph properties are normalized to one leading pPr because baseline export generated duplicate nodes that triggered native PowerPoint repair. Original packages and dependencies are never rewritten.

`mountStaticSlide(container, document, slideId)` creates an immutable validated read-only slide with the same internal DOM renderer, returning `{element, ready, destroy}`. `ready` waits for font/image decoding. Hosts must load the fork stylesheet and explicit font bytes. Matching viewport placement matters for exact rotated-shape antialiasing. This API is a bounded proof, not an accepted production renderer.

47 Node regressions pass. The sibling SlideVids `proofs/slidekit` Docker/browser harness checks actual Notes/Arrange interactions, JSON/PPTX reopen, exact editor/static pixels and Remotion PNG frames. H.264 with PNG intermediates gives max channel delta 7 against trial criterion 4, so the strict runner remains nonzero. No production dependency, migration, audio or workflow integration occurred. Local changes are not pushed. Native simple Notes export opens without repair and displays Unicode/paragraphs; native save is blocked by inactive Office licensing. Customer decks remain untested.

## Controlled host opening and Notes history

The integration fork exposes `ui.open: false` in createEditor options. This hides the native Open control and prevents its internal opening method from replacing a host-owned deck. Export and manual editing remain available. The app owns reviewed one-time creation and revision restore. Repeated identical slide patches, including duplicate Notes input events, do not add redundant undo history. Fork version 0.2.0-sv.3 includes this regression fix. The T02 static renderer artifact 0.2.0-sv.1 remains frozen.
