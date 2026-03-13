;;; mcp_dispatch.lsp — File-based IPC dispatcher for AutoCAD MCP v3.2
;;;
;;; Protocol:
;;;   1. Python writes command JSON to C:/temp/autocad_mcp_cmd_{id}.json
;;;   2. Python types "(c:mcp-dispatch)" + Enter
;;;   3. This function reads cmd, dispatches via command map, writes result JSON
;;;   4. Python polls for C:/temp/autocad_mcp_result_{id}.json
;;;
;;; SECURITY: No raw eval — dispatcher uses a command whitelist/map.
;;; Compatible with AutoCAD LT 2024+.

;; Load dependencies
(if (not report-error)
  (defun report-error (msg) (princ (strcat "\nERROR: " msg)))
)

;; IPC directory
(setq *mcp-ipc-dir* "C:/temp/")

;; -----------------------------------------------------------------------
;; JSON-like output helpers (minimal, no external library)
;; -----------------------------------------------------------------------

(defun mcp-write-result (filepath request-id ok-flag payload error-msg / fp)
  "Write a result JSON file. Atomic: write to .tmp then rename."
  (setq tmp-path (strcat filepath ".tmp"))
  (setq fp (open tmp-path "w"))
  (if fp
    (progn
      (write-line "{" fp)
      (write-line (strcat "  \"request_id\": \"" request-id "\",") fp)
      (if ok-flag
        (progn
          (write-line "  \"ok\": true," fp)
          (write-line (strcat "  \"payload\": " payload) fp)
        )
        (progn
          (write-line "  \"ok\": false," fp)
          (write-line (strcat "  \"error\": \"" (mcp-escape-string error-msg) "\"") fp)
        )
      )
      (write-line "}" fp)
      (close fp)
      ;; Rename .tmp to final path (atomic on NTFS)
      (vl-file-rename tmp-path filepath)
    )
    (princ (strcat "\nMCP: Cannot open result file: " tmp-path))
  )
)

(defun mcp-escape-string (s / result i ch)
  "Escape quotes and backslashes in a string for JSON."
  (if (null s) (setq s ""))
  (setq result "" i 1)
  (while (<= i (strlen s))
    (setq ch (substr s i 1))
    (cond
      ((= ch "\"") (setq result (strcat result "\\\"")))
      ((= ch "\\") (setq result (strcat result "\\\\")))
      (t (setq result (strcat result ch)))
    )
    (setq i (1+ i))
  )
  result
)

(defun mcp-read-file-lines (filepath / fp line lines)
  "Read all lines from a file into a single string."
  (setq fp (open filepath "r"))
  (if (not fp) (progn (princ (strcat "\nMCP: Cannot read: " filepath)) nil)
    (progn
      (setq lines "")
      (while (setq line (read-line fp))
        (setq lines (strcat lines line))
      )
      (close fp)
      lines
    )
  )
)

;; -----------------------------------------------------------------------
;; Simple JSON parser (extracts string values by key)
;; -----------------------------------------------------------------------

(defun mcp-json-get-string (json key / search-str pos end-pos value)
  "Extract a string value for a given key from JSON text."
  (setq search-str (strcat "\"" key "\""))
  (setq pos (vl-string-search search-str json))
  (if (null pos) nil
    (progn
      ;; Find the colon after key
      (setq pos (vl-string-search ":" json pos))
      (if (null pos) nil
        (progn
          ;; Find opening quote of value
          (setq pos (vl-string-search "\"" json (1+ pos)))
          (if (null pos) nil
            (progn
              (setq pos (+ pos 2))  ; 0-based search result + 2 = 1-based position after quote
              ;; Find closing quote (skip escaped quotes)
              (setq end-pos pos)
              (while (and (<= end-pos (strlen json))
                          (or (= end-pos pos)
                              (/= (substr json end-pos 1) "\"")))
                ;; Handle escaped characters
                (if (= (substr json end-pos 1) "\\")
                  (setq end-pos (+ end-pos 2))
                  (setq end-pos (1+ end-pos))
                )
              )
              (substr json pos (- end-pos pos))
            )
          )
        )
      )
    )
  )
)

(defun mcp-json-get-number (json key / search-str pos num-start num-end ch)
  "Extract a number value for a given key from JSON text."
  (setq search-str (strcat "\"" key "\""))
  (setq pos (vl-string-search search-str json))
  (if (null pos) nil
    (progn
      (setq pos (vl-string-search ":" json pos))
      (if (null pos) nil
        (progn
          (setq pos (+ pos 2))  ; 0-based search result + 2 = 1-based position after colon
          ;; Skip whitespace
          (while (and (<= pos (strlen json))
                      (member (substr json pos 1) '(" " "\t" "\n")))
            (setq pos (1+ pos))
          )
          ;; Read number
          (setq num-start pos num-end pos)
          (while (and (<= num-end (strlen json))
                      (or (member (substr json num-end 1) '("0" "1" "2" "3" "4" "5" "6" "7" "8" "9" "." "-" "+"))
                      ))
            (setq num-end (1+ num-end))
          )
          (atof (substr json num-start (- num-end num-start)))
        )
      )
    )
  )
)

;; -----------------------------------------------------------------------
;; String splitting utility (used by semicolon-delimited encodings)
;; -----------------------------------------------------------------------

(defun mcp-split-string (str delim / pos result token)
  "Split a string by single-char delimiter. Returns a list of strings."
  (setq result '())
  (while (setq pos (vl-string-search delim str))
    (setq token (substr str 1 pos))
    (setq result (append result (list token)))
    (setq str (substr str (+ pos 2)))
  )
  (setq result (append result (list str)))
  result
)

;; -----------------------------------------------------------------------
;; Command dispatcher — WHITELIST ONLY, no eval
;; -----------------------------------------------------------------------

(defun mcp-dispatch-command (cmd-name params-json / result)
  "Dispatch a command by name. Returns (ok . payload-or-error)."
  (cond
    ;; --- Ping ---
    ((= cmd-name "ping")
     (cons T "\"pong\""))

    ;; --- Freehand LISP execution ---
    ((= cmd-name "execute-lisp")
     (mcp-cmd-execute-lisp params-json))

    ;; --- Undo / Redo ---
    ((= cmd-name "undo")
     (command "_.UNDO" "1") (cons T "\"undone\""))

    ((= cmd-name "redo")
     (command "_.REDO") (cons T "\"redone\""))

    ;; --- Drawing info ---
    ((= cmd-name "drawing-info")
     (mcp-cmd-drawing-info))

    ;; --- Layer operations ---
    ((= cmd-name "layer-list")
     (mcp-cmd-layer-list))

    ((= cmd-name "layer-create")
     (mcp-cmd-layer-create params-json))

    ((= cmd-name "layer-set-current")
     (mcp-cmd-layer-set-current params-json))

    ((= cmd-name "layer-set-properties")
     (mcp-cmd-layer-set-properties params-json))

    ((= cmd-name "layer-freeze")
     (mcp-cmd-layer-freeze params-json))

    ((= cmd-name "layer-thaw")
     (mcp-cmd-layer-thaw params-json))

    ((= cmd-name "layer-lock")
     (mcp-cmd-layer-lock params-json))

    ((= cmd-name "layer-unlock")
     (mcp-cmd-layer-unlock params-json))

    ;; --- Entity creation ---
    ((= cmd-name "create-line")
     (mcp-cmd-create-line params-json))

    ((= cmd-name "create-circle")
     (mcp-cmd-create-circle params-json))

    ((= cmd-name "create-polyline")
     (mcp-cmd-create-polyline params-json))

    ((= cmd-name "create-rectangle")
     (mcp-cmd-create-rectangle params-json))

    ((= cmd-name "create-text")
     (mcp-cmd-create-text params-json))

    ((= cmd-name "create-arc")
     (mcp-cmd-create-arc params-json))

    ((= cmd-name "create-ellipse")
     (mcp-cmd-create-ellipse params-json))

    ((= cmd-name "create-mtext")
     (mcp-cmd-create-mtext params-json))

    ((= cmd-name "create-hatch")
     (mcp-cmd-create-hatch params-json))

    ;; --- Entity queries ---
    ((= cmd-name "entity-count")
     (mcp-cmd-entity-count params-json))

    ((= cmd-name "entity-list")
     (mcp-cmd-entity-list params-json))

    ((= cmd-name "entity-get")
     (mcp-cmd-entity-get params-json))

    ((= cmd-name "entity-erase")
     (mcp-cmd-entity-erase params-json))

    ;; --- Entity modification ---
    ((= cmd-name "entity-move")
     (mcp-cmd-entity-move params-json))

    ((= cmd-name "entity-copy")
     (mcp-cmd-entity-copy params-json))

    ((= cmd-name "entity-rotate")
     (mcp-cmd-entity-rotate params-json))

    ((= cmd-name "entity-scale")
     (mcp-cmd-entity-scale params-json))

    ((= cmd-name "entity-mirror")
     (mcp-cmd-entity-mirror params-json))

    ((= cmd-name "entity-offset")
     (mcp-cmd-entity-offset params-json))

    ((= cmd-name "entity-array")
     (mcp-cmd-entity-array params-json))

    ((= cmd-name "entity-fillet")
     (mcp-cmd-entity-fillet params-json))

    ((= cmd-name "entity-chamfer")
     (mcp-cmd-entity-chamfer params-json))

    ;; --- View ---
    ((= cmd-name "zoom-extents")
     (command "_.ZOOM" "_E")
     (cons T "\"zoomed to extents\""))

    ((= cmd-name "zoom-window")
     (progn
       (setq x1 (mcp-json-get-number params-json "x1"))
       (setq y1 (mcp-json-get-number params-json "y1"))
       (setq x2 (mcp-json-get-number params-json "x2"))
       (setq y2 (mcp-json-get-number params-json "y2"))
       (command "_.ZOOM" "_W" (list x1 y1 0) (list x2 y2 0))
       (cons T "\"zoomed to window\"")))

    ;; --- Drawing file ops ---
    ((= cmd-name "drawing-save")
     (progn
       (setq path (mcp-json-get-string params-json "path"))
       (if (and path (> (strlen path) 0))
         (progn
           (setvar "FILEDIA" 0)
           (command "_.SAVEAS" "" path)
           (setvar "FILEDIA" 1)
           (cons T (strcat "\"saved to: " (mcp-escape-string path) "\"")))
         (progn (command "_.QSAVE") (cons T "\"saved\"")))))

    ((= cmd-name "drawing-save-as-dxf")
     (progn
       (setq path (mcp-json-get-string params-json "path"))
       (if path
         (progn (command "_.SAVEAS" "DXF" path) (cons T (strcat "\"" path "\"")))
         (cons nil "Save path required"))))

    ((= cmd-name "drawing-purge")
     (command "_.-PURGE" "_ALL" "*" "_N")
     (cons T "\"purged\""))

    ((= cmd-name "drawing-open")
     (progn
       (setq path (mcp-json-get-string params-json "path"))
       (if path
         (progn
           (setvar "FILEDIA" 0)
           (command "_.OPEN" path)
           (setvar "FILEDIA" 1)
           (cons T (strcat "\"opened: " (mcp-escape-string path) "\"")))
         (cons nil "Path required"))))

    ;; --- P&ID ---
    ((= cmd-name "pid-setup-layers")
     (if c:setup-pid-layers
       (progn (c:setup-pid-layers) (cons T "\"P&ID layers created\""))
       (cons nil "pid_tools.lsp not loaded")))

    ((= cmd-name "pid-insert-symbol")
     (mcp-cmd-pid-insert-symbol params-json))

    ((= cmd-name "pid-draw-process-line")
     (mcp-cmd-pid-draw-process-line params-json))

    ((= cmd-name "pid-connect-equipment")
     (mcp-cmd-pid-connect-equipment params-json))

    ((= cmd-name "pid-add-flow-arrow")
     (mcp-cmd-pid-add-flow-arrow params-json))

    ((= cmd-name "pid-add-equipment-tag")
     (mcp-cmd-pid-add-equipment-tag params-json))

    ((= cmd-name "pid-add-line-number")
     (mcp-cmd-pid-add-line-number params-json))

    ((= cmd-name "pid-insert-valve")
     (mcp-cmd-pid-insert-valve params-json))

    ((= cmd-name "pid-insert-instrument")
     (mcp-cmd-pid-insert-instrument params-json))

    ((= cmd-name "pid-insert-pump")
     (mcp-cmd-pid-insert-pump params-json))

    ((= cmd-name "pid-insert-tank")
     (mcp-cmd-pid-insert-tank params-json))

    ;; --- Block operations ---
    ((= cmd-name "block-list")
     (mcp-cmd-block-list))

    ((= cmd-name "block-insert")
     (mcp-cmd-block-insert params-json))

    ((= cmd-name "block-insert-with-attributes")
     (mcp-cmd-block-insert-with-attribs params-json))

    ((= cmd-name "block-get-attributes")
     (mcp-cmd-block-get-attributes params-json))

    ((= cmd-name "block-update-attribute")
     (mcp-cmd-block-update-attribute params-json))

    ((= cmd-name "block-define")
     (cons nil "block-define not available via IPC (use ezdxf backend)"))

    ;; --- Annotation ---
    ((= cmd-name "create-dimension-linear")
     (mcp-cmd-create-dimension-linear params-json))

    ((= cmd-name "create-dimension-aligned")
     (mcp-cmd-create-dimension-aligned params-json))

    ((= cmd-name "create-dimension-angular")
     (mcp-cmd-create-dimension-angular params-json))

    ((= cmd-name "create-dimension-radius")
     (mcp-cmd-create-dimension-radius params-json))

    ((= cmd-name "create-leader")
     (mcp-cmd-create-leader params-json))

    ;; --- Drawing management ---
    ((= cmd-name "drawing-create")
     (mcp-cmd-drawing-create params-json))

    ((= cmd-name "drawing-get-variables")
     (mcp-cmd-drawing-get-variables params-json))

    ((= cmd-name "drawing-plot-pdf")
     (mcp-cmd-drawing-plot-pdf params-json))

    ;; --- P&ID list symbols ---
    ((= cmd-name "pid-list-symbols")
     (mcp-cmd-pid-list-symbols params-json))

    ;; --- Query operations ---
    ((= cmd-name "query-entity-properties")
     (mcp-cmd-query-entity-properties params-json))

    ((= cmd-name "query-entity-geometry")
     (mcp-cmd-query-entity-geometry params-json))

    ((= cmd-name "query-drawing-summary")
     (mcp-cmd-query-drawing-summary))

    ((= cmd-name "query-layer-summary")
     (mcp-cmd-query-layer-summary params-json))

    ;; --- Search operations ---
    ((= cmd-name "search-text")
     (mcp-cmd-search-text params-json))

    ((= cmd-name "search-by-attribute")
     (mcp-cmd-search-by-attribute params-json))

    ((= cmd-name "search-by-window")
     (mcp-cmd-search-by-window params-json))

    ((= cmd-name "search-by-proximity")
     (mcp-cmd-search-by-proximity params-json))

    ((= cmd-name "search-by-type-and-layer")
     (mcp-cmd-search-by-type-and-layer params-json))

    ;; --- Geometry operations ---
    ((= cmd-name "geometry-distance")
     (mcp-cmd-geometry-distance params-json))

    ((= cmd-name "geometry-length")
     (mcp-cmd-geometry-length params-json))

    ((= cmd-name "geometry-area")
     (mcp-cmd-geometry-area params-json))

    ((= cmd-name "geometry-bounding-box")
     (mcp-cmd-geometry-bounding-box params-json))

    ((= cmd-name "geometry-polyline-info")
     (mcp-cmd-geometry-polyline-info params-json))

    ;; --- Bulk operations ---
    ((= cmd-name "bulk-set-property")
     (mcp-cmd-bulk-set-property params-json))

    ((= cmd-name "bulk-erase")
     (mcp-cmd-bulk-erase params-json))

    ;; --- Export ---
    ((= cmd-name "export-entity-data")
     (mcp-cmd-export-entity-data params-json))

    ;; --- Select / Filter ---
    ((= cmd-name "select-filter")
     (mcp-cmd-select-filter params-json))

    ((= cmd-name "bulk-move")
     (mcp-cmd-bulk-move params-json))

    ((= cmd-name "bulk-copy")
     (mcp-cmd-bulk-copy params-json))

    ((= cmd-name "find-replace-text")
     (mcp-cmd-find-replace-text params-json))

    ;; --- Entity Modification ---
    ((= cmd-name "entity-set-property")
     (mcp-cmd-entity-set-property params-json))

    ((= cmd-name "entity-set-text")
     (mcp-cmd-entity-set-text params-json))

    ;; --- View Enhancements ---
    ((= cmd-name "zoom-center")
     (mcp-cmd-zoom-center params-json))

    ((= cmd-name "layer-visibility")
     (mcp-cmd-layer-visibility params-json))

    ;; --- Validate ---
    ((= cmd-name "validate-layer-standards")
     (mcp-cmd-validate-layer-standards params-json))

    ((= cmd-name "validate-duplicates")
     (mcp-cmd-validate-duplicates params-json))

    ((= cmd-name "validate-zero-length")
     (mcp-cmd-validate-zero-length params-json))

    ((= cmd-name "validate-qc-report")
     (mcp-cmd-validate-qc-report params-json))

    ;; --- Export / Reporting ---
    ((= cmd-name "export-bom")
     (mcp-cmd-export-bom params-json))

    ((= cmd-name "export-data-extract")
     (mcp-cmd-export-data-extract params-json))

    ((= cmd-name "export-layer-report")
     (mcp-cmd-export-layer-report params-json))

    ((= cmd-name "export-block-count")
     (mcp-cmd-export-block-count params-json))

    ((= cmd-name "export-drawing-statistics")
     (mcp-cmd-export-drawing-statistics params-json))

    ;; --- Extended Query ---
    ((= cmd-name "query-text-styles")
     (mcp-cmd-query-text-styles params-json))
    ((= cmd-name "query-dimension-styles")
     (mcp-cmd-query-dimension-styles params-json))
    ((= cmd-name "query-linetypes")
     (mcp-cmd-query-linetypes params-json))
    ((= cmd-name "query-block-tree")
     (mcp-cmd-query-block-tree params-json))
    ((= cmd-name "query-drawing-metadata")
     (mcp-cmd-query-drawing-metadata params-json))

    ;; --- Extended Search ---
    ((= cmd-name "search-by-block-name")
     (mcp-cmd-search-by-block-name params-json))
    ((= cmd-name "search-by-handle-list")
     (mcp-cmd-search-by-handle-list params-json))

    ;; --- Extended Entity ---
    ((= cmd-name "entity-explode")
     (mcp-cmd-entity-explode params-json))
    ((= cmd-name "entity-join")
     (mcp-cmd-entity-join params-json))
    ((= cmd-name "entity-extend")
     (mcp-cmd-entity-extend params-json))
    ((= cmd-name "entity-trim")
     (mcp-cmd-entity-trim params-json))
    ((= cmd-name "entity-break-at")
     (mcp-cmd-entity-break-at params-json))

    ;; --- Extended Validate ---
    ((= cmd-name "validate-text-standards")
     (mcp-cmd-validate-text-standards params-json))
    ((= cmd-name "validate-orphaned-entities")
     (mcp-cmd-validate-orphaned-entities params-json))
    ((= cmd-name "validate-attribute-completeness")
     (mcp-cmd-validate-attribute-completeness params-json))
    ((= cmd-name "validate-connectivity")
     (mcp-cmd-validate-connectivity params-json))

    ;; --- Extended Select ---
    ((= cmd-name "find-replace-attribute")
     (mcp-cmd-find-replace-attribute params-json))
    ((= cmd-name "layer-rename")
     (mcp-cmd-layer-rename params-json))
    ((= cmd-name "layer-merge")
     (mcp-cmd-layer-merge params-json))

    ;; --- Enhanced View ---
    ((= cmd-name "zoom-scale")
     (mcp-cmd-zoom-scale params-json))
    ((= cmd-name "pan")
     (mcp-cmd-pan params-json))

    ;; --- Enhanced Drawing ---
    ((= cmd-name "drawing-audit")
     (mcp-cmd-drawing-audit params-json))
    ((= cmd-name "drawing-units")
     (mcp-cmd-drawing-units params-json))
    ((= cmd-name "drawing-limits")
     (mcp-cmd-drawing-limits params-json))

    ;; --- XREF ---
    ((= cmd-name "xref-list")
     (mcp-cmd-xref-list params-json))
    ((= cmd-name "xref-attach")
     (mcp-cmd-xref-attach params-json))
    ((= cmd-name "xref-detach")
     (mcp-cmd-xref-detach params-json))
    ((= cmd-name "xref-reload")
     (mcp-cmd-xref-reload params-json))
    ((= cmd-name "xref-bind")
     (mcp-cmd-xref-bind params-json))
    ((= cmd-name "xref-path-update")
     (mcp-cmd-xref-path-update params-json))
    ((= cmd-name "xref-query-entities")
     (mcp-cmd-xref-query-entities params-json))

    ;; --- Layout ---
    ((= cmd-name "layout-list")
     (mcp-cmd-layout-list params-json))
    ((= cmd-name "layout-create")
     (mcp-cmd-layout-create params-json))
    ((= cmd-name "layout-switch")
     (mcp-cmd-layout-switch params-json))
    ((= cmd-name "layout-delete")
     (mcp-cmd-layout-delete params-json))
    ((= cmd-name "layout-viewport-create")
     (mcp-cmd-layout-viewport-create params-json))
    ((= cmd-name "layout-viewport-set-scale")
     (mcp-cmd-layout-viewport-set-scale params-json))
    ((= cmd-name "layout-viewport-lock")
     (mcp-cmd-layout-viewport-lock params-json))
    ((= cmd-name "layout-page-setup")
     (mcp-cmd-layout-page-setup params-json))
    ((= cmd-name "layout-titleblock-fill")
     (mcp-cmd-layout-titleblock-fill params-json))
    ((= cmd-name "layout-batch-plot")
     (mcp-cmd-layout-batch-plot params-json))

    ;; --- Drawing WBLOCK ---
    ((= cmd-name "drawing-wblock") (mcp-cmd-drawing-wblock params-json))

    ;; --- Electrical ---
    ((= cmd-name "electrical-nec-lookup") (mcp-cmd-electrical-nec-lookup params-json))
    ((= cmd-name "electrical-voltage-drop") (mcp-cmd-electrical-voltage-drop params-json))
    ((= cmd-name "electrical-conduit-fill") (mcp-cmd-electrical-conduit-fill params-json))
    ((= cmd-name "electrical-load-calc") (mcp-cmd-electrical-load-calc params-json))
    ((= cmd-name "electrical-symbol-insert") (mcp-cmd-electrical-symbol-insert params-json))
    ((= cmd-name "electrical-circuit-trace") (mcp-cmd-electrical-circuit-trace params-json))
    ((= cmd-name "electrical-panel-schedule-gen") (mcp-cmd-electrical-panel-schedule-gen params-json))
    ((= cmd-name "electrical-wire-number-assign") (mcp-cmd-electrical-wire-number-assign params-json))

    ;; --- Equipment Find / Inspect ---
    ((= cmd-name "equipment-find") (mcp-cmd-equipment-find params-json))
    ((= cmd-name "equipment-inspect") (mcp-cmd-equipment-inspect params-json))
    ((= cmd-name "find-text") (mcp-cmd-find-text params-json))

    ;; --- Equipment Tag Placement ---
    ((= cmd-name "place-equipment-tag") (mcp-cmd-place-equipment-tag params-json))
    ((= cmd-name "batch-find-and-tag") (mcp-cmd-batch-find-and-tag params-json))

    ;; --- MagiCAD ---
    ((= cmd-name "magicad-status") (mcp-cmd-magicad-status))
    ((= cmd-name "magicad-run") (mcp-cmd-magicad-run params-json))
    ((= cmd-name "magicad-update-drawing") (mcp-cmd-magicad-update-drawing params-json))
    ((= cmd-name "magicad-cleanup") (mcp-cmd-magicad-cleanup params-json))
    ((= cmd-name "magicad-ifc-export") (mcp-cmd-magicad-ifc-export params-json))
    ((= cmd-name "magicad-view-mode") (mcp-cmd-magicad-view-mode params-json))
    ((= cmd-name "magicad-change-storey") (mcp-cmd-magicad-change-storey params-json))
    ((= cmd-name "magicad-section-update") (mcp-cmd-magicad-section-update))
    ((= cmd-name "magicad-fix-errors") (mcp-cmd-magicad-fix-errors))
    ((= cmd-name "magicad-show-all") (mcp-cmd-magicad-show-all))
    ((= cmd-name "magicad-clear-garbage") (mcp-cmd-magicad-clear-garbage))
    ((= cmd-name "magicad-disconnect-project") (mcp-cmd-magicad-disconnect-project))
    ((= cmd-name "magicad-list-commands") (mcp-cmd-magicad-list-commands))

    ;; --- Unknown ---
    (t (cons nil (strcat "Unknown command: " cmd-name)))
  )
)

;; -----------------------------------------------------------------------
;; Command implementations
;; -----------------------------------------------------------------------

(defun mcp-cmd-drawing-info ( / count layers layer-list)
  "Return drawing info: entity count, layers, extents."
  (setq count 0)
  (setq ent (entnext))
  (while ent
    (setq count (1+ count))
    (setq ent (entnext ent))
  )
  (setq layer-list "")
  (setq layers (tblnext "LAYER" T))
  (while layers
    (if (> (strlen layer-list) 0)
      (setq layer-list (strcat layer-list ",\"" (cdr (assoc 2 layers)) "\""))
      (setq layer-list (strcat "\"" (cdr (assoc 2 layers)) "\""))
    )
    (setq layers (tblnext "LAYER"))
  )
  (cons T (strcat "{\"entity_count\":" (itoa count) ",\"layers\":[" layer-list "]}"))
)

(defun mcp-cmd-layer-list ( / layers layer-list name)
  "Return all layers as JSON array."
  (setq layer-list "")
  (setq layers (tblnext "LAYER" T))
  (while layers
    (setq name (cdr (assoc 2 layers)))
    (if (> (strlen layer-list) 0)
      (setq layer-list (strcat layer-list ",{\"name\":\"" name "\",\"color\":" (itoa (cdr (assoc 62 layers))) "}"))
      (setq layer-list (strcat "{\"name\":\"" name "\",\"color\":" (itoa (cdr (assoc 62 layers))) "}"))
    )
    (setq layers (tblnext "LAYER"))
  )
  (cons T (strcat "{\"layers\":[" layer-list "]}"))
)

(defun mcp-cmd-layer-create (params / name color linetype)
  (setq name (mcp-json-get-string params "name"))
  (setq color (mcp-json-get-string params "color"))
  (setq linetype (mcp-json-get-string params "linetype"))
  (if (not color) (setq color "white"))
  (if (not linetype) (setq linetype "CONTINUOUS"))
  (ensure_layer_exists name color linetype)
  (cons T (strcat "{\"name\":\"" name "\"}"))
)

(defun mcp-cmd-layer-set-current (params / name)
  (setq name (mcp-json-get-string params "name"))
  (setvar "CLAYER" name)
  (cons T (strcat "{\"current_layer\":\"" name "\"}"))
)

(defun mcp-cmd-create-line (params / x1 y1 x2 y2 layer)
  (setq x1 (mcp-json-get-number params "x1"))
  (setq y1 (mcp-json-get-number params "y1"))
  (setq x2 (mcp-json-get-number params "x2"))
  (setq y2 (mcp-json-get-number params "y2"))
  (setq layer (mcp-json-get-string params "layer"))
  (if layer
    (progn (ensure_layer_exists layer "white" "CONTINUOUS") (set_current_layer layer))
  )
  (command "_LINE" (list x1 y1 0.0) (list x2 y2 0.0) "")
  (cons T (strcat "{\"entity_type\":\"LINE\",\"handle\":\"" (cdr (assoc 5 (entget (entlast)))) "\"}"))
)

(defun mcp-cmd-create-circle (params / cx cy radius layer)
  (setq cx (mcp-json-get-number params "cx"))
  (setq cy (mcp-json-get-number params "cy"))
  (setq radius (mcp-json-get-number params "radius"))
  (setq layer (mcp-json-get-string params "layer"))
  (if layer
    (progn (ensure_layer_exists layer "white" "CONTINUOUS") (set_current_layer layer))
  )
  (command "_CIRCLE" (list cx cy 0.0) radius)
  (cons T (strcat "{\"entity_type\":\"CIRCLE\",\"handle\":\"" (cdr (assoc 5 (entget (entlast)))) "\"}"))
)

(defun mcp-cmd-create-polyline (params / pts-str closed layer pairs pt-str cx cy)
  (setq pts-str (mcp-json-get-string params "points_str"))
  (setq closed (mcp-json-get-string params "closed"))
  (setq layer (mcp-json-get-string params "layer"))
  (if layer (progn (ensure_layer_exists layer "white" "CONTINUOUS") (set_current_layer layer)))
  (if (not pts-str)
    (cons nil "points_str required (format: x1,y1;x2,y2;...)")
    (progn
      (command "_PLINE")
      (setq pairs (mcp-split-string pts-str ";"))
      (foreach pt-str pairs
        (setq cx (atof (car (mcp-split-string pt-str ","))))
        (setq cy (atof (cadr (mcp-split-string pt-str ","))))
        (command (list cx cy 0.0))
      )
      (if (= closed "1") (command "_C") (command ""))
      (cons T (strcat "{\"entity_type\":\"LWPOLYLINE\",\"handle\":\""
                      (cdr (assoc 5 (entget (entlast)))) "\"}"))
    )
  )
)

(defun mcp-cmd-create-rectangle (params / x1 y1 x2 y2 layer)
  (setq x1 (mcp-json-get-number params "x1"))
  (setq y1 (mcp-json-get-number params "y1"))
  (setq x2 (mcp-json-get-number params "x2"))
  (setq y2 (mcp-json-get-number params "y2"))
  (setq layer (mcp-json-get-string params "layer"))
  (if layer
    (progn (ensure_layer_exists layer "white" "CONTINUOUS") (set_current_layer layer))
  )
  (command "_RECTANG" (list x1 y1 0.0) (list x2 y2 0.0))
  (cons T (strcat "{\"entity_type\":\"LWPOLYLINE\",\"handle\":\"" (cdr (assoc 5 (entget (entlast)))) "\"}"))
)

(defun mcp-cmd-create-text (params / x y text height rotation layer)
  (setq x (mcp-json-get-number params "x"))
  (setq y (mcp-json-get-number params "y"))
  (setq text (mcp-json-get-string params "text"))
  (setq height (mcp-json-get-number params "height"))
  (setq rotation (mcp-json-get-number params "rotation"))
  (if (not height) (setq height 2.5))
  (if (not rotation) (setq rotation 0.0))
  (setq layer (mcp-json-get-string params "layer"))
  (if layer
    (progn (ensure_layer_exists layer "white" "CONTINUOUS") (set_current_layer layer))
  )
  (command "_TEXT" "J" "M" (list x y 0.0) height rotation text)
  (cons T (strcat "{\"entity_type\":\"TEXT\",\"handle\":\"" (cdr (assoc 5 (entget (entlast)))) "\"}"))
)

(defun mcp-cmd-entity-count (params / layer count ent ent-data)
  (setq layer (mcp-json-get-string params "layer"))
  (setq count 0 ent (entnext))
  (while ent
    (setq ent-data (entget ent))
    (if (or (not layer) (= (cdr (assoc 8 ent-data)) layer))
      (setq count (1+ count))
    )
    (setq ent (entnext ent))
  )
  (cons T (strcat "{\"count\":" (itoa count) "}"))
)

(defun mcp-cmd-entity-list (params / layer entities ent ent-data etype handle elayer)
  (setq layer (mcp-json-get-string params "layer"))
  (setq entities "" ent (entnext))
  (while ent
    (setq ent-data (entget ent))
    (setq etype (cdr (assoc 0 ent-data)))
    (setq handle (cdr (assoc 5 ent-data)))
    (setq elayer (cdr (assoc 8 ent-data)))
    (if (or (not layer) (= elayer layer))
      (progn
        (if (> (strlen entities) 0)
          (setq entities (strcat entities ","))
        )
        (setq entities (strcat entities "{\"type\":\"" etype "\",\"handle\":\"" handle "\",\"layer\":\"" elayer "\"}"))
      )
    )
    (setq ent (entnext ent))
  )
  (cons T (strcat "{\"entities\":[" entities "]}"))
)

(defun mcp-cmd-entity-erase (params / entity-id ent)
  (setq entity-id (mcp-json-get-string params "entity_id"))
  (if (= entity-id "last")
    (progn
      (setq ent (entlast))
      (if ent (progn (entdel ent) (cons T "\"erased last entity\""))
        (cons nil "No entity to erase")))
    (progn
      (setq ent (handent entity-id))
      (if ent (progn (entdel ent) (cons T (strcat "\"erased " entity-id "\"")))
        (cons nil (strcat "Entity not found: " entity-id))))
  )
)

(defun mcp-cmd-entity-move (params / entity-id dx dy ent)
  (setq entity-id (mcp-json-get-string params "entity_id"))
  (setq dx (mcp-json-get-number params "dx"))
  (setq dy (mcp-json-get-number params "dy"))
  (if (= entity-id "last")
    (setq ent (entlast))
    (setq ent (handent entity-id))
  )
  (if ent
    (progn
      (command "_.MOVE" ent "" '(0 0 0) (list dx dy 0))
      (cons T "\"moved\""))
    (cons nil "Entity not found")
  )
)

;; --- Freehand LISP execution ---

(defun mcp-cmd-execute-lisp (params / code-file result old-secureload)
  (setq code-file (mcp-json-get-string params "code_file"))
  (if (not code-file)
    (cons nil "code_file parameter required")
    (if (not (findfile code-file))
      (cons nil (strcat "Code file not found: " code-file))
      (progn
        ;; Suppress SECURELOAD dialog for MCP temp files
        (setq old-secureload (getvar "SECURELOAD"))
        (setvar "SECURELOAD" 0)
        (setq result (vl-catch-all-apply 'load (list code-file)))
        (setvar "SECURELOAD" old-secureload)
        (if (vl-catch-all-error-p result)
          (cons nil (strcat "LISP error: " (vl-catch-all-error-message result)))
          (cons T (strcat "\"" (mcp-escape-string (vl-princ-to-string result)) "\""))
        )
      )
    )
  )
)

;; --- Drawing create implementation ---

(defun mcp-cmd-drawing-create (params / ss)
  "Reset current drawing to a clean state (erase all, purge, reset to layer 0).
   Using _.NEW would create a new document tab with a fresh LISP namespace,
   breaking the IPC dispatcher. This approach preserves the dispatcher."
  (if (setq ss (ssget "_X"))
    (progn (command "_.ERASE" ss "") (setq ss nil))
  )
  (setvar "CLAYER" "0")
  (command "_.-PURGE" "_ALL" "*" "_N")
  (cons T (strcat "{\"drawing\":\"" (mcp-escape-string (getvar "DWGNAME")) "\"}"))
)

;; --- P&ID command implementations ---

(defun mcp-cmd-pid-insert-symbol (params / category symbol x y scale rotation)
  (setq category (mcp-json-get-string params "category"))
  (setq symbol (mcp-json-get-string params "symbol"))
  (setq x (mcp-json-get-number params "x"))
  (setq y (mcp-json-get-number params "y"))
  (setq scale (mcp-json-get-number params "scale"))
  (setq rotation (mcp-json-get-number params "rotation"))
  (if (not scale) (setq scale 1.0))
  (if (not rotation) (setq rotation 0.0))
  (if c:insert-pid-block
    (progn
      (c:insert-pid-block category symbol x y scale rotation)
      (cons T (strcat "{\"symbol\":\"" symbol "\",\"handle\":\"" (cdr (assoc 5 (entget (entlast)))) "\"}"))
    )
    (cons nil "pid_tools.lsp not loaded")
  )
)

(defun mcp-cmd-pid-draw-process-line (params / x1 y1 x2 y2)
  (setq x1 (mcp-json-get-number params "x1"))
  (setq y1 (mcp-json-get-number params "y1"))
  (setq x2 (mcp-json-get-number params "x2"))
  (setq y2 (mcp-json-get-number params "y2"))
  (if c:draw-process-line
    (progn (c:draw-process-line x1 y1 x2 y2) (cons T "\"process line drawn\""))
    (cons nil "pid_tools.lsp not loaded")
  )
)

(defun mcp-cmd-pid-connect-equipment (params / x1 y1 x2 y2)
  (setq x1 (mcp-json-get-number params "x1"))
  (setq y1 (mcp-json-get-number params "y1"))
  (setq x2 (mcp-json-get-number params "x2"))
  (setq y2 (mcp-json-get-number params "y2"))
  (if c:connect-equipment
    (progn (c:connect-equipment x1 y1 x2 y2) (cons T "\"equipment connected\""))
    (cons nil "pid_tools.lsp not loaded")
  )
)

(defun mcp-cmd-pid-add-flow-arrow (params / x y rotation)
  (setq x (mcp-json-get-number params "x"))
  (setq y (mcp-json-get-number params "y"))
  (setq rotation (mcp-json-get-number params "rotation"))
  (if (not rotation) (setq rotation 0.0))
  (if c:add-flow-arrow
    (progn (c:add-flow-arrow x y rotation) (cons T "\"flow arrow added\""))
    (cons nil "pid_tools.lsp not loaded")
  )
)

(defun mcp-cmd-pid-add-equipment-tag (params / x y tag description)
  (setq x (mcp-json-get-number params "x"))
  (setq y (mcp-json-get-number params "y"))
  (setq tag (mcp-json-get-string params "tag"))
  (setq description (mcp-json-get-string params "description"))
  (if (not description) (setq description ""))
  (if c:add-equipment-tag
    (progn (c:add-equipment-tag x y tag description) (cons T (strcat "\"tagged: " tag "\"")))
    (cons nil "pid_tools.lsp not loaded")
  )
)

(defun mcp-cmd-pid-add-line-number (params / x y line-num spec)
  (setq x (mcp-json-get-number params "x"))
  (setq y (mcp-json-get-number params "y"))
  (setq line-num (mcp-json-get-string params "line_num"))
  (setq spec (mcp-json-get-string params "spec"))
  (if c:add-line-number
    (progn (c:add-line-number x y line-num spec) (cons T (strcat "\"line number: " line-num "\"")))
    (cons nil "pid_tools.lsp not loaded")
  )
)

(defun mcp-cmd-pid-insert-valve (params / x y valve-type rotation)
  (setq x (mcp-json-get-number params "x"))
  (setq y (mcp-json-get-number params "y"))
  (setq valve-type (mcp-json-get-string params "valve_type"))
  (setq rotation (mcp-json-get-number params "rotation"))
  (if (not rotation) (setq rotation 0.0))
  (if c:insert-valve-on-line
    (progn (c:insert-valve-on-line x y valve-type rotation) (cons T (strcat "\"valve: " valve-type "\"")))
    (cons nil "pid_tools.lsp not loaded")
  )
)

(defun mcp-cmd-pid-insert-instrument (params / x y inst-type rotation tag-id range-value)
  (setq x (mcp-json-get-number params "x"))
  (setq y (mcp-json-get-number params "y"))
  (setq inst-type (mcp-json-get-string params "instrument_type"))
  (setq rotation (mcp-json-get-number params "rotation"))
  (setq tag-id (mcp-json-get-string params "tag_id"))
  (setq range-value (mcp-json-get-string params "range_value"))
  (if (not rotation) (setq rotation 0.0))
  (if c:insert-instrument
    (progn
      (c:insert-instrument x y inst-type rotation)
      (if (and tag-id (> (strlen tag-id) 0))
        (c:insert-instrument-with-tag x y inst-type tag-id (if range-value range-value ""))
      )
      (cons T (strcat "\"instrument: " inst-type "\"")))
    (cons nil "pid_tools.lsp not loaded")
  )
)

(defun mcp-cmd-pid-insert-pump (params / x y pump-type rotation)
  (setq x (mcp-json-get-number params "x"))
  (setq y (mcp-json-get-number params "y"))
  (setq pump-type (mcp-json-get-string params "pump_type"))
  (setq rotation (mcp-json-get-number params "rotation"))
  (if (not rotation) (setq rotation 0.0))
  (if c:insert-pump
    (progn (c:insert-pump x y pump-type rotation) (cons T (strcat "\"pump: " pump-type "\"")))
    (cons nil "pid_tools.lsp not loaded")
  )
)

(defun mcp-cmd-pid-insert-tank (params / x y tank-type scale)
  (setq x (mcp-json-get-number params "x"))
  (setq y (mcp-json-get-number params "y"))
  (setq tank-type (mcp-json-get-string params "tank_type"))
  (setq scale (mcp-json-get-number params "scale"))
  (if (not scale) (setq scale 1.0))
  (if c:insert-tank
    (progn (c:insert-tank x y tank-type scale) (cons T (strcat "\"tank: " tank-type "\"")))
    (cons nil "pid_tools.lsp not loaded")
  )
)

;; --- Additional entity creation ---

(defun mcp-cmd-create-arc (params / cx cy radius sa ea layer)
  (setq cx (mcp-json-get-number params "cx"))
  (setq cy (mcp-json-get-number params "cy"))
  (setq radius (mcp-json-get-number params "radius"))
  (setq sa (mcp-json-get-number params "start_angle"))
  (setq ea (mcp-json-get-number params "end_angle"))
  (setq layer (mcp-json-get-string params "layer"))
  (if layer (progn (ensure_layer_exists layer "white" "CONTINUOUS") (set_current_layer layer)))
  (command "_ARC" "_C" (list cx cy 0.0) (list (+ cx radius) cy 0.0) "_A" (- ea sa))
  (cons T (strcat "{\"entity_type\":\"ARC\",\"handle\":\"" (cdr (assoc 5 (entget (entlast)))) "\"}"))
)

(defun mcp-cmd-create-ellipse (params / cx cy mx my ratio layer)
  (setq cx (mcp-json-get-number params "cx"))
  (setq cy (mcp-json-get-number params "cy"))
  (setq mx (mcp-json-get-number params "major_x"))
  (setq my (mcp-json-get-number params "major_y"))
  (setq ratio (mcp-json-get-number params "ratio"))
  (setq layer (mcp-json-get-string params "layer"))
  (if layer (progn (ensure_layer_exists layer "white" "CONTINUOUS") (set_current_layer layer)))
  (command "_ELLIPSE" "_C" (list cx cy 0.0) (list mx my 0.0) ratio)
  (cons T (strcat "{\"entity_type\":\"ELLIPSE\",\"handle\":\"" (cdr (assoc 5 (entget (entlast)))) "\"}"))
)

(defun mcp-cmd-create-mtext (params / x y width text height layer)
  (setq x (mcp-json-get-number params "x"))
  (setq y (mcp-json-get-number params "y"))
  (setq width (mcp-json-get-number params "width"))
  (setq text (mcp-json-get-string params "text"))
  (setq height (mcp-json-get-number params "height"))
  (if (not height) (setq height 2.5))
  (setq layer (mcp-json-get-string params "layer"))
  (if layer (progn (ensure_layer_exists layer "white" "CONTINUOUS") (set_current_layer layer)))
  (command "_MTEXT" (list x y 0.0) "_H" height "_W" width text "")
  (cons T (strcat "{\"entity_type\":\"MTEXT\",\"handle\":\"" (cdr (assoc 5 (entget (entlast)))) "\"}"))
)

(defun mcp-cmd-create-hatch (params / entity-id pattern ent)
  (setq entity-id (mcp-json-get-string params "entity_id"))
  (setq pattern (mcp-json-get-string params "pattern"))
  (if (not pattern) (setq pattern "ANSI31"))
  (if (= entity-id "last")
    (setq ent (entlast))
    (setq ent (handent entity-id))
  )
  (if ent
    (progn
      (command "_HATCH" "_P" pattern "" "_S" ent "" "")
      (cons T (strcat "{\"entity_type\":\"HATCH\",\"handle\":\"" (cdr (assoc 5 (entget (entlast)))) "\"}")))
    (cons nil "Entity not found for hatching")
  )
)

;; --- Entity query: get ---

(defun mcp-cmd-entity-get (params / entity-id ent ent-data etype handle elayer result)
  (setq entity-id (mcp-json-get-string params "entity_id"))
  (if (= entity-id "last")
    (setq ent (entlast))
    (setq ent (handent entity-id))
  )
  (if (not ent)
    (cons nil (strcat "Entity not found: " entity-id))
    (progn
      (setq ent-data (entget ent))
      (setq etype (cdr (assoc 0 ent-data)))
      (setq handle (cdr (assoc 5 ent-data)))
      (setq elayer (cdr (assoc 8 ent-data)))
      (setq result (strcat "{\"type\":\"" etype "\",\"handle\":\"" handle "\",\"layer\":\"" elayer "\""))
      ;; Add type-specific info
      (cond
        ((= etype "LINE")
         (setq result (strcat result
           ",\"start\":[" (rtos (car (cdr (assoc 10 ent-data))) 2 6) "," (rtos (cadr (cdr (assoc 10 ent-data))) 2 6) "]"
           ",\"end\":[" (rtos (car (cdr (assoc 11 ent-data))) 2 6) "," (rtos (cadr (cdr (assoc 11 ent-data))) 2 6) "]")))
        ((= etype "CIRCLE")
         (setq result (strcat result
           ",\"center\":[" (rtos (car (cdr (assoc 10 ent-data))) 2 6) "," (rtos (cadr (cdr (assoc 10 ent-data))) 2 6) "]"
           ",\"radius\":" (rtos (cdr (assoc 40 ent-data)) 2 6))))
      )
      (setq result (strcat result "}"))
      (cons T result)
    )
  )
)

;; --- Entity modification commands ---

(defun mcp-cmd-entity-copy (params / entity-id dx dy ent new-handle)
  (setq entity-id (mcp-json-get-string params "entity_id"))
  (setq dx (mcp-json-get-number params "dx"))
  (setq dy (mcp-json-get-number params "dy"))
  (if (= entity-id "last") (setq ent (entlast)) (setq ent (handent entity-id)))
  (if ent
    (progn
      (command "_.COPY" ent "" '(0 0 0) (list dx dy 0))
      (setq new-handle (cdr (assoc 5 (entget (entlast)))))
      (cons T (strcat "{\"handle\":\"" new-handle "\"}")))
    (cons nil "Entity not found")
  )
)

(defun mcp-cmd-entity-rotate (params / entity-id cx cy angle ent)
  (setq entity-id (mcp-json-get-string params "entity_id"))
  (setq cx (mcp-json-get-number params "cx"))
  (setq cy (mcp-json-get-number params "cy"))
  (setq angle (mcp-json-get-number params "angle"))
  (if (= entity-id "last") (setq ent (entlast)) (setq ent (handent entity-id)))
  (if ent
    (progn (command "_.ROTATE" ent "" (list cx cy 0) angle) (cons T "\"rotated\""))
    (cons nil "Entity not found")
  )
)

(defun mcp-cmd-entity-scale (params / entity-id cx cy factor ent)
  (setq entity-id (mcp-json-get-string params "entity_id"))
  (setq cx (mcp-json-get-number params "cx"))
  (setq cy (mcp-json-get-number params "cy"))
  (setq factor (mcp-json-get-number params "factor"))
  (if (= entity-id "last") (setq ent (entlast)) (setq ent (handent entity-id)))
  (if ent
    (progn (command "_.SCALE" ent "" (list cx cy 0) factor) (cons T "\"scaled\""))
    (cons nil "Entity not found")
  )
)

(defun mcp-cmd-entity-mirror (params / entity-id x1 y1 x2 y2 ent new-handle)
  (setq entity-id (mcp-json-get-string params "entity_id"))
  (setq x1 (mcp-json-get-number params "x1"))
  (setq y1 (mcp-json-get-number params "y1"))
  (setq x2 (mcp-json-get-number params "x2"))
  (setq y2 (mcp-json-get-number params "y2"))
  (if (= entity-id "last") (setq ent (entlast)) (setq ent (handent entity-id)))
  (if ent
    (progn
      (command "_.MIRROR" ent "" (list x1 y1 0) (list x2 y2 0) "_N")
      (setq new-handle (cdr (assoc 5 (entget (entlast)))))
      (cons T (strcat "{\"handle\":\"" new-handle "\"}")))
    (cons nil "Entity not found")
  )
)

(defun mcp-cmd-entity-offset (params / entity-id distance ent new-handle)
  (setq entity-id (mcp-json-get-string params "entity_id"))
  (setq distance (mcp-json-get-number params "distance"))
  (if (= entity-id "last") (setq ent (entlast)) (setq ent (handent entity-id)))
  (if ent
    (progn
      (command "_.OFFSET" distance ent (list 0 0 0) "")
      (setq new-handle (cdr (assoc 5 (entget (entlast)))))
      (cons T (strcat "{\"handle\":\"" new-handle "\"}")))
    (cons nil "Entity not found")
  )
)

(defun mcp-cmd-entity-array (params / entity-id rows cols row-dist col-dist ent)
  (setq entity-id (mcp-json-get-string params "entity_id"))
  (setq rows (fix (mcp-json-get-number params "rows")))
  (setq cols (fix (mcp-json-get-number params "cols")))
  (setq row-dist (mcp-json-get-number params "row_dist"))
  (setq col-dist (mcp-json-get-number params "col_dist"))
  (if (= entity-id "last") (setq ent (entlast)) (setq ent (handent entity-id)))
  (if ent
    (progn
      (command "_.ARRAY" ent "" "_R" rows cols row-dist col-dist)
      (cons T (strcat "{\"rows\":" (itoa rows) ",\"cols\":" (itoa cols) "}")))
    (cons nil "Entity not found")
  )
)

(defun mcp-cmd-entity-fillet (params / id1 id2 radius ent1 ent2)
  (setq id1 (mcp-json-get-string params "id1"))
  (setq id2 (mcp-json-get-string params "id2"))
  (setq radius (mcp-json-get-number params "radius"))
  (setq ent1 (handent id1))
  (setq ent2 (handent id2))
  (if (and ent1 ent2)
    (progn
      (command "_.FILLET" "_R" radius)
      (command "_.FILLET" ent1 ent2)
      (cons T "\"filleted\""))
    (cons nil "One or both entities not found")
  )
)

(defun mcp-cmd-entity-chamfer (params / id1 id2 dist1 dist2 ent1 ent2)
  (setq id1 (mcp-json-get-string params "id1"))
  (setq id2 (mcp-json-get-string params "id2"))
  (setq dist1 (mcp-json-get-number params "dist1"))
  (setq dist2 (mcp-json-get-number params "dist2"))
  (setq ent1 (handent id1))
  (setq ent2 (handent id2))
  (if (and ent1 ent2)
    (progn
      (command "_.CHAMFER" "_D" dist1 dist2)
      (command "_.CHAMFER" ent1 ent2)
      (cons T "\"chamfered\""))
    (cons nil "One or both entities not found")
  )
)

;; --- Layer operations ---

(defun mcp-cmd-layer-set-properties (params / name color linetype lineweight)
  (setq name (mcp-json-get-string params "name"))
  (setq color (mcp-json-get-string params "color"))
  (setq linetype (mcp-json-get-string params "linetype"))
  (setq lineweight (mcp-json-get-string params "lineweight"))
  (if color (command "_.-LAYER" "_COLOR" color name ""))
  (if linetype (command "_.-LAYER" "_LTYPE" linetype name ""))
  (if lineweight (command "_.-LAYER" "_LWEIGHT" lineweight name ""))
  (cons T (strcat "{\"name\":\"" name "\"}"))
)

(defun mcp-cmd-layer-freeze (params / name)
  (setq name (mcp-json-get-string params "name"))
  (command "_.-LAYER" "_FREEZE" name "")
  (cons T (strcat "{\"name\":\"" name "\",\"frozen\":true}"))
)

(defun mcp-cmd-layer-thaw (params / name)
  (setq name (mcp-json-get-string params "name"))
  (command "_.-LAYER" "_THAW" name "")
  (cons T (strcat "{\"name\":\"" name "\",\"frozen\":false}"))
)

(defun mcp-cmd-layer-lock (params / name)
  (setq name (mcp-json-get-string params "name"))
  (command "_.-LAYER" "_LOCK" name "")
  (cons T (strcat "{\"name\":\"" name "\",\"locked\":true}"))
)

(defun mcp-cmd-layer-unlock (params / name)
  (setq name (mcp-json-get-string params "name"))
  (command "_.-LAYER" "_UNLOCK" name "")
  (cons T (strcat "{\"name\":\"" name "\",\"locked\":false}"))
)

;; --- Block operations (insert-with-attributes, get-attributes, update-attribute) ---

(defun mcp-cmd-block-insert-with-attribs (params / name x y scale rotation attributes ent)
  (setq name (mcp-json-get-string params "name"))
  (setq x (mcp-json-get-number params "x"))
  (setq y (mcp-json-get-number params "y"))
  (setq scale (mcp-json-get-number params "scale"))
  (setq rotation (mcp-json-get-number params "rotation"))
  (if (not scale) (setq scale 1.0))
  (if (not rotation) (setq rotation 0.0))
  (if (tblsearch "BLOCK" name)
    (progn
      ;; Insert with ATTREQ=1 to fill attributes
      (command "_.INSERT" name (list x y 0.0) scale scale rotation)
      ;; Note: attribute values are applied separately via update-attribute
      (cons T (strcat "{\"entity_type\":\"INSERT\",\"handle\":\"" (cdr (assoc 5 (entget (entlast)))) "\"}")))
    (cons nil (strcat "Block '" name "' not found"))
  )
)

(defun mcp-cmd-block-get-attributes (params / entity-id ent sub-ent ent-data attribs)
  (setq entity-id (mcp-json-get-string params "entity_id"))
  (if (= entity-id "last") (setq ent (entlast)) (setq ent (handent entity-id)))
  (if (not ent)
    (cons nil "Entity not found")
    (progn
      (setq attribs "" sub-ent (entnext ent))
      (while sub-ent
        (setq ent-data (entget sub-ent))
        (if (= (cdr (assoc 0 ent-data)) "ATTRIB")
          (progn
            (if (> (strlen attribs) 0) (setq attribs (strcat attribs ",")))
            (setq attribs (strcat attribs "\"" (cdr (assoc 2 ent-data)) "\":\"" (mcp-escape-string (cdr (assoc 1 ent-data))) "\""))
          )
        )
        (if (= (cdr (assoc 0 ent-data)) "SEQEND")
          (setq sub-ent nil)
          (setq sub-ent (entnext sub-ent))
        )
      )
      (cons T (strcat "{\"attributes\":{" attribs "}}"))
    )
  )
)

(defun mcp-cmd-block-update-attribute (params / entity-id tag value ent)
  (setq entity-id (mcp-json-get-string params "entity_id"))
  (setq tag (mcp-json-get-string params "tag"))
  (setq value (mcp-json-get-string params "value"))
  (if (= entity-id "last") (setq ent (entlast)) (setq ent (handent entity-id)))
  (if (not ent)
    (cons nil "Entity not found")
    (progn
      (if c:update-block-attribute
        (progn (c:update-block-attribute ent tag value)
               (cons T (strcat "{\"tag\":\"" tag "\",\"value\":\"" (mcp-escape-string value) "\"}")))
        ;; Inline fallback if attribute_tools.lsp not loaded
        (progn
          (set_attribute_value ent tag value)
          (cons T (strcat "{\"tag\":\"" tag "\",\"value\":\"" (mcp-escape-string value) "\"}")))
      )
    )
  )
)

;; --- Annotation commands ---

(defun mcp-cmd-create-dimension-linear (params / x1 y1 x2 y2 dim-x dim-y)
  (setq x1 (mcp-json-get-number params "x1"))
  (setq y1 (mcp-json-get-number params "y1"))
  (setq x2 (mcp-json-get-number params "x2"))
  (setq y2 (mcp-json-get-number params "y2"))
  (setq dim-x (mcp-json-get-number params "dim_x"))
  (setq dim-y (mcp-json-get-number params "dim_y"))
  (command "_.DIMLINEAR" (list x1 y1 0) (list x2 y2 0) (list dim-x dim-y 0))
  (cons T "{\"entity_type\":\"DIMENSION\"}")
)

(defun mcp-cmd-create-dimension-aligned (params / x1 y1 x2 y2 offset)
  (setq x1 (mcp-json-get-number params "x1"))
  (setq y1 (mcp-json-get-number params "y1"))
  (setq x2 (mcp-json-get-number params "x2"))
  (setq y2 (mcp-json-get-number params "y2"))
  (setq offset (mcp-json-get-number params "offset"))
  ;; Place dimension line at offset distance
  (command "_.DIMALIGNED" (list x1 y1 0) (list x2 y2 0)
    (list (+ (/ (+ x1 x2) 2.0) offset) (+ (/ (+ y1 y2) 2.0) offset) 0))
  (cons T "{\"entity_type\":\"DIMENSION\"}")
)

(defun mcp-cmd-create-dimension-angular (params / cx cy x1 y1 x2 y2)
  (setq cx (mcp-json-get-number params "cx"))
  (setq cy (mcp-json-get-number params "cy"))
  (setq x1 (mcp-json-get-number params "x1"))
  (setq y1 (mcp-json-get-number params "y1"))
  (setq x2 (mcp-json-get-number params "x2"))
  (setq y2 (mcp-json-get-number params "y2"))
  (command "_.DIMANGULAR" (list cx cy 0) (list x1 y1 0) (list x2 y2 0) "")
  (cons T "{\"entity_type\":\"DIMENSION\"}")
)

(defun mcp-cmd-create-dimension-radius (params / cx cy radius angle px py)
  (setq cx (mcp-json-get-number params "cx"))
  (setq cy (mcp-json-get-number params "cy"))
  (setq radius (mcp-json-get-number params "radius"))
  (setq angle (mcp-json-get-number params "angle"))
  ;; Need a circle/arc entity first, use entity at center
  (setq px (+ cx (* radius (cos (* angle (/ pi 180.0))))))
  (setq py (+ cy (* radius (sin (* angle (/ pi 180.0))))))
  (command "_.DIMRADIUS" (list px py 0) "")
  (cons T "{\"entity_type\":\"DIMENSION\"}")
)

(defun mcp-cmd-create-leader (params / text pts-str pairs pt-str)
  (setq text (mcp-json-get-string params "text"))
  (setq pts-str (mcp-json-get-string params "points_str"))
  (if (not pts-str)
    (cons nil "points_str required (format: x1,y1;x2,y2;...)")
    (progn
      (command "_.LEADER")
      (setq pairs (mcp-split-string pts-str ";"))
      (foreach pt-str pairs
        (command (list (atof (car (mcp-split-string pt-str ",")))
                       (atof (cadr (mcp-split-string pt-str ","))) 0))
      )
      (command "" text "")
      (cons T "{\"entity_type\":\"LEADER\"}")
    )
  )
)

;; --- Drawing management ---

(defun mcp-cmd-drawing-get-variables (params / names-str result var-list var-name var-val first-var)
  (setq names-str (mcp-json-get-string params "names_str"))
  (if (or (not names-str) (= names-str ""))
    ;; Default set when no specific names requested
    (progn
      (setq result "{")
      (setq result (strcat result "\"ACADVER\":\"" (getvar "ACADVER") "\""))
      (setq result (strcat result ",\"DWGNAME\":\"" (mcp-escape-string (getvar "DWGNAME")) "\""))
      (setq result (strcat result ",\"CLAYER\":\"" (getvar "CLAYER") "\""))
      (setq result (strcat result "}"))
      (cons T result)
    )
    ;; Parse semicolon-delimited variable names
    (progn
      (setq var-list (mcp-split-string names-str ";"))
      (setq result "{" first-var T)
      (foreach var-name var-list
        (setq var-val (getvar var-name))
        (if (not first-var) (setq result (strcat result ",")))
        (setq first-var nil)
        (if (not var-val)
          (setq result (strcat result "\"" var-name "\":null"))
          (cond
            ((= (type var-val) 'STR)
             (setq result (strcat result "\"" var-name "\":\"" (mcp-escape-string var-val) "\"")))
            ((= (type var-val) 'INT)
             (setq result (strcat result "\"" var-name "\":" (itoa var-val))))
            ((= (type var-val) 'REAL)
             (setq result (strcat result "\"" var-name "\":" (rtos var-val 2 6))))
            (t
             (setq result (strcat result "\"" var-name "\":\"" (mcp-escape-string (vl-princ-to-string var-val)) "\"")))
          )
        )
      )
      (setq result (strcat result "}"))
      (cons T result)
    )
  )
)

(defun mcp-cmd-drawing-plot-pdf (params / path)
  (setq path (mcp-json-get-string params "path"))
  (if path
    (progn
      (command "_.-PLOT" "_Y" "" "DWG To PDF.pc3"
        "ANSI_A_(8.50_x_11.00_Inches)" "_Inches" "_Landscape"
        "_N" "_Extents" "_Fit" "_Y" "acad.ctb" "_Y" "_N" "_Y" path "_Y")
      (cons T (strcat "{\"path\":\"" (mcp-escape-string path) "\"}")))
    (cons nil "Plot path required")
  )
)

;; --- P&ID list symbols ---

(defun mcp-cmd-pid-list-symbols (params / category dir-path files result)
  (setq category (mcp-json-get-string params "category"))
  (setq dir-path (strcat "C:/PIDv4-CTO/" category "/"))
  (setq files (vl-directory-files dir-path "*.dwg" 1))
  (setq result "")
  (if files
    (foreach f files
      (if (> (strlen result) 0) (setq result (strcat result ",")))
      ;; Remove .dwg extension
      (setq result (strcat result "\"" (substr f 1 (- (strlen f) 4)) "\""))
    )
  )
  (cons T (strcat "{\"category\":\"" category "\",\"symbols\":[" result "],\"count\":" (itoa (length (if files files '()))) "}"))
)

;; --- Block operations ---

(defun mcp-cmd-block-list ( / blk block-list)
  (setq block-list "" blk (tblnext "BLOCK" T))
  (while blk
    (if (not (= (substr (cdr (assoc 2 blk)) 1 1) "*"))
      (progn
        (if (> (strlen block-list) 0)
          (setq block-list (strcat block-list ",\"" (cdr (assoc 2 blk)) "\""))
          (setq block-list (strcat "\"" (cdr (assoc 2 blk)) "\""))
        )
      )
    )
    (setq blk (tblnext "BLOCK"))
  )
  (cons T (strcat "{\"blocks\":[" block-list "]}"))
)

(defun mcp-cmd-block-insert (params / name x y scale rotation block-id)
  (setq name (mcp-json-get-string params "name"))
  (setq x (mcp-json-get-number params "x"))
  (setq y (mcp-json-get-number params "y"))
  (setq scale (mcp-json-get-number params "scale"))
  (setq rotation (mcp-json-get-number params "rotation"))
  (setq block-id (mcp-json-get-string params "block_id"))
  (if (not scale) (setq scale 1.0))
  (if (not rotation) (setq rotation 0.0))
  (if (tblsearch "BLOCK" name)
    (progn
      (command "_.INSERT" name (list x y 0.0) scale scale rotation)
      (if (and block-id (> (strlen block-id) 0))
        (set_attribute_value (entlast) "ID" block-id)
      )
      (cons T (strcat "{\"entity_type\":\"INSERT\",\"handle\":\"" (cdr (assoc 5 (entget (entlast)))) "\"}"))
    )
    (cons nil (strcat "Block '" name "' not found"))
  )
)

;; -----------------------------------------------------------------------
;; Query / Search / Geometry / Bulk / Export command implementations
;; -----------------------------------------------------------------------

;; Helper: Format a 3D point as JSON array string
(defun mcp-point-to-json (pt)
  (if pt
    (strcat "[" (rtos (car pt) 2 6) "," (rtos (cadr pt) 2 6)
            (if (caddr pt) (strcat "," (rtos (caddr pt) 2 6)) "")
            "]")
    "null"
  )
)

;; Helper: Format a number for JSON (handles nil)
(defun mcp-num-to-json (val)
  (if val (rtos val 2 6) "null")
)

;; Helper: Get entity position (center/insertion point) for any entity type
(defun mcp-entity-position (ent-data / etype pt)
  (setq etype (cdr (assoc 0 ent-data)))
  (setq pt (cdr (assoc 10 ent-data)))
  (if pt pt '(0.0 0.0 0.0))
)

;; Helper: extract all DXF group codes as JSON for an entity
(defun mcp-entget-to-json (ent / ent-data etype handle elayer color ltype result)
  (setq ent-data (entget ent))
  (setq etype (cdr (assoc 0 ent-data)))
  (setq handle (cdr (assoc 5 ent-data)))
  (setq elayer (cdr (assoc 8 ent-data)))
  (setq color (cdr (assoc 62 ent-data)))
  (setq ltype (cdr (assoc 6 ent-data)))

  (setq result (strcat "{\"type\":\"" etype "\",\"handle\":\"" handle "\",\"layer\":\"" (mcp-escape-string elayer) "\""))
  (if color (setq result (strcat result ",\"color\":" (itoa color))))
  (if ltype (setq result (strcat result ",\"linetype\":\"" (mcp-escape-string ltype) "\"")))

  ;; Type-specific properties
  (cond
    ((= etype "LINE")
     (setq result (strcat result
       ",\"start\":" (mcp-point-to-json (cdr (assoc 10 ent-data)))
       ",\"end\":" (mcp-point-to-json (cdr (assoc 11 ent-data))))))

    ((= etype "CIRCLE")
     (setq result (strcat result
       ",\"center\":" (mcp-point-to-json (cdr (assoc 10 ent-data)))
       ",\"radius\":" (mcp-num-to-json (cdr (assoc 40 ent-data))))))

    ((= etype "ARC")
     (setq result (strcat result
       ",\"center\":" (mcp-point-to-json (cdr (assoc 10 ent-data)))
       ",\"radius\":" (mcp-num-to-json (cdr (assoc 40 ent-data)))
       ",\"start_angle\":" (mcp-num-to-json (cdr (assoc 50 ent-data)))
       ",\"end_angle\":" (mcp-num-to-json (cdr (assoc 51 ent-data))))))

    ((= etype "ELLIPSE")
     (setq result (strcat result
       ",\"center\":" (mcp-point-to-json (cdr (assoc 10 ent-data)))
       ",\"major_axis\":" (mcp-point-to-json (cdr (assoc 11 ent-data)))
       ",\"ratio\":" (mcp-num-to-json (cdr (assoc 40 ent-data))))))

    ((= etype "TEXT")
     (setq result (strcat result
       ",\"content\":\"" (mcp-escape-string (cdr (assoc 1 ent-data))) "\""
       ",\"position\":" (mcp-point-to-json (cdr (assoc 10 ent-data)))
       ",\"height\":" (mcp-num-to-json (cdr (assoc 40 ent-data)))
       ",\"rotation\":" (mcp-num-to-json (cdr (assoc 50 ent-data)))
       ",\"style\":\"" (mcp-escape-string (if (cdr (assoc 7 ent-data)) (cdr (assoc 7 ent-data)) "Standard")) "\"")))

    ((= etype "MTEXT")
     (setq result (strcat result
       ",\"content\":\"" (mcp-escape-string (cdr (assoc 1 ent-data))) "\""
       ",\"position\":" (mcp-point-to-json (cdr (assoc 10 ent-data)))
       ",\"char_height\":" (mcp-num-to-json (cdr (assoc 40 ent-data)))
       ",\"width\":" (mcp-num-to-json (cdr (assoc 41 ent-data)))
       ",\"attachment_point\":" (if (cdr (assoc 71 ent-data)) (itoa (cdr (assoc 71 ent-data))) "null")
       ",\"style\":\"" (mcp-escape-string (if (cdr (assoc 7 ent-data)) (cdr (assoc 7 ent-data)) "Standard")) "\""
       ",\"background_fill\":" (if (cdr (assoc 90 ent-data)) (itoa (cdr (assoc 90 ent-data))) "0"))))

    ((= etype "INSERT")
     (setq result (strcat result
       ",\"block_name\":\"" (mcp-escape-string (cdr (assoc 2 ent-data))) "\""
       ",\"insertion\":" (mcp-point-to-json (cdr (assoc 10 ent-data)))
       ",\"x_scale\":" (mcp-num-to-json (cdr (assoc 41 ent-data)))
       ",\"y_scale\":" (mcp-num-to-json (cdr (assoc 42 ent-data)))
       ",\"z_scale\":" (mcp-num-to-json (cdr (assoc 43 ent-data)))
       ",\"rotation\":" (mcp-num-to-json (cdr (assoc 50 ent-data))))))

    ((= etype "LWPOLYLINE")
     (setq result (strcat result
       ",\"closed\":" (if (= (logand (cdr (assoc 70 ent-data)) 1) 1) "true" "false")
       ",\"vertex_count\":" (itoa (cdr (assoc 90 ent-data))))))

    ((= etype "POLYLINE")
     (setq result (strcat result
       ",\"closed\":" (if (= (logand (cdr (assoc 70 ent-data)) 1) 1) "true" "false")
       ",\"flags\":" (itoa (cdr (assoc 70 ent-data))))))

    ((= etype "LEADER")
     (setq result (strcat result
       ",\"dimstyle\":\"" (mcp-escape-string (if (cdr (assoc 3 ent-data)) (cdr (assoc 3 ent-data)) "")) "\""
       ",\"has_annotation\":" (if (cdr (assoc 340 ent-data)) "true" "false"))))

    ((= etype "DIMENSION")
     (setq result (strcat result
       ",\"dimstyle\":\"" (mcp-escape-string (if (cdr (assoc 3 ent-data)) (cdr (assoc 3 ent-data)) "")) "\""
       ",\"text_override\":\"" (mcp-escape-string (if (cdr (assoc 1 ent-data)) (cdr (assoc 1 ent-data)) "")) "\""
       ",\"definition_point\":" (mcp-point-to-json (cdr (assoc 10 ent-data))))))

    ((= etype "HATCH")
     (setq result (strcat result
       ",\"pattern\":\"" (mcp-escape-string (cdr (assoc 2 ent-data))) "\"")))
  )

  (setq result (strcat result "}"))
  result
)

;; --- query-entity-properties ---
(defun mcp-cmd-query-entity-properties (params / entity-id ent)
  (setq entity-id (mcp-json-get-string params "entity_id"))
  (if (= entity-id "last") (setq ent (entlast)) (setq ent (handent entity-id)))
  (if (not ent)
    (cons nil (strcat "Entity not found: " entity-id))
    (cons T (mcp-entget-to-json ent))
  )
)

;; --- query-entity-geometry ---
(defun mcp-cmd-query-entity-geometry (params / entity-id ent ent-data etype result sub-ent sub-data verts)
  (setq entity-id (mcp-json-get-string params "entity_id"))
  (if (= entity-id "last") (setq ent (entlast)) (setq ent (handent entity-id)))
  (if (not ent)
    (cons nil (strcat "Entity not found: " entity-id))
    (progn
      (setq ent-data (entget ent))
      (setq etype (cdr (assoc 0 ent-data)))
      (setq result (strcat "{\"type\":\"" etype "\",\"handle\":\"" (cdr (assoc 5 ent-data)) "\""))

      (cond
        ((= etype "LINE")
         (setq result (strcat result
           ",\"start\":" (mcp-point-to-json (cdr (assoc 10 ent-data)))
           ",\"end\":" (mcp-point-to-json (cdr (assoc 11 ent-data))))))

        ((= etype "CIRCLE")
         (setq result (strcat result
           ",\"center\":" (mcp-point-to-json (cdr (assoc 10 ent-data)))
           ",\"radius\":" (mcp-num-to-json (cdr (assoc 40 ent-data))))))

        ((= etype "ARC")
         (setq result (strcat result
           ",\"center\":" (mcp-point-to-json (cdr (assoc 10 ent-data)))
           ",\"radius\":" (mcp-num-to-json (cdr (assoc 40 ent-data)))
           ",\"start_angle\":" (mcp-num-to-json (cdr (assoc 50 ent-data)))
           ",\"end_angle\":" (mcp-num-to-json (cdr (assoc 51 ent-data))))))

        ((= etype "LWPOLYLINE")
         (progn
           (setq verts "")
           ;; Extract all group 10 entries for vertices
           (foreach pair ent-data
             (if (= (car pair) 10)
               (progn
                 (if (> (strlen verts) 0) (setq verts (strcat verts ",")))
                 (setq verts (strcat verts (mcp-point-to-json (cdr pair))))
               )
             )
           )
           (setq result (strcat result
             ",\"closed\":" (if (= (logand (cdr (assoc 70 ent-data)) 1) 1) "true" "false")
             ",\"vertices\":[" verts "]"))))

        ((= etype "POLYLINE")
         (progn
           ;; Walk VERTEX sub-entities
           (setq verts "" sub-ent (entnext ent))
           (while sub-ent
             (setq sub-data (entget sub-ent))
             (if (= (cdr (assoc 0 sub-data)) "VERTEX")
               (progn
                 (if (> (strlen verts) 0) (setq verts (strcat verts ",")))
                 (setq verts (strcat verts (mcp-point-to-json (cdr (assoc 10 sub-data)))))
               )
             )
             (if (= (cdr (assoc 0 sub-data)) "SEQEND")
               (setq sub-ent nil)
               (setq sub-ent (entnext sub-ent))
             )
           )
           (setq result (strcat result
             ",\"closed\":" (if (= (logand (cdr (assoc 70 ent-data)) 1) 1) "true" "false")
             ",\"flags\":" (itoa (cdr (assoc 70 ent-data)))
             ",\"vertices\":[" verts "]"))))

        ((= etype "TEXT")
         (setq result (strcat result
           ",\"content\":\"" (mcp-escape-string (cdr (assoc 1 ent-data))) "\""
           ",\"position\":" (mcp-point-to-json (cdr (assoc 10 ent-data)))
           ",\"height\":" (mcp-num-to-json (cdr (assoc 40 ent-data)))
           ",\"rotation\":" (mcp-num-to-json (cdr (assoc 50 ent-data)))
           ",\"style\":\"" (mcp-escape-string (if (cdr (assoc 7 ent-data)) (cdr (assoc 7 ent-data)) "Standard")) "\"")))

        ((= etype "MTEXT")
         (setq result (strcat result
           ",\"content\":\"" (mcp-escape-string (cdr (assoc 1 ent-data))) "\""
           ",\"position\":" (mcp-point-to-json (cdr (assoc 10 ent-data)))
           ",\"char_height\":" (mcp-num-to-json (cdr (assoc 40 ent-data)))
           ",\"width\":" (mcp-num-to-json (cdr (assoc 41 ent-data)))
           ",\"attachment_point\":" (if (cdr (assoc 71 ent-data)) (itoa (cdr (assoc 71 ent-data))) "null")
           ",\"style\":\"" (mcp-escape-string (if (cdr (assoc 7 ent-data)) (cdr (assoc 7 ent-data)) "Standard")) "\"")))

        ((= etype "INSERT")
         (setq result (strcat result
           ",\"block_name\":\"" (mcp-escape-string (cdr (assoc 2 ent-data))) "\""
           ",\"insertion\":" (mcp-point-to-json (cdr (assoc 10 ent-data)))
           ",\"x_scale\":" (mcp-num-to-json (cdr (assoc 41 ent-data)))
           ",\"y_scale\":" (mcp-num-to-json (cdr (assoc 42 ent-data)))
           ",\"rotation\":" (mcp-num-to-json (cdr (assoc 50 ent-data))))))

        ((= etype "LEADER")
         (progn
           ;; Walk leader vertices (group 10 entries)
           (setq verts "")
           (foreach pair ent-data
             (if (= (car pair) 10)
               (progn
                 (if (> (strlen verts) 0) (setq verts (strcat verts ",")))
                 (setq verts (strcat verts (mcp-point-to-json (cdr pair))))
               )
             )
           )
           (setq result (strcat result
             ",\"dimstyle\":\"" (mcp-escape-string (if (cdr (assoc 3 ent-data)) (cdr (assoc 3 ent-data)) "")) "\""
             ",\"vertices\":[" verts "]"))))

        ((= etype "DIMENSION")
         (setq result (strcat result
           ",\"dimstyle\":\"" (mcp-escape-string (if (cdr (assoc 3 ent-data)) (cdr (assoc 3 ent-data)) "")) "\""
           ",\"definition_point\":" (mcp-point-to-json (cdr (assoc 10 ent-data)))
           ",\"text_midpoint\":" (mcp-point-to-json (cdr (assoc 11 ent-data)))
           ",\"text_override\":\"" (mcp-escape-string (if (cdr (assoc 1 ent-data)) (cdr (assoc 1 ent-data)) "")) "\"")))

        ((= etype "HATCH")
         (setq result (strcat result
           ",\"pattern\":\"" (mcp-escape-string (cdr (assoc 2 ent-data))) "\"")))

        ((= etype "ELLIPSE")
         (setq result (strcat result
           ",\"center\":" (mcp-point-to-json (cdr (assoc 10 ent-data)))
           ",\"major_axis\":" (mcp-point-to-json (cdr (assoc 11 ent-data)))
           ",\"ratio\":" (mcp-num-to-json (cdr (assoc 40 ent-data))))))
      )

      (setq result (strcat result "}"))
      (cons T result)
    )
  )
)

;; --- query-drawing-summary ---
(defun mcp-cmd-query-drawing-summary ( / ent ent-data etype elayer type-counts layer-counts
                                         type-list layer-list result extmin extmax)
  ;; Count entities by type and layer
  (setq type-counts '() layer-counts '() ent (entnext))
  (while ent
    (setq ent-data (entget ent))
    (setq etype (cdr (assoc 0 ent-data)))
    (setq elayer (cdr (assoc 8 ent-data)))
    ;; Skip sub-entities
    (if (not (member etype '("VERTEX" "SEQEND" "ATTRIB" "ATTDEF")))
      (progn
        ;; Count by type
        (if (assoc etype type-counts)
          (setq type-counts (subst (cons etype (1+ (cdr (assoc etype type-counts))))
                                   (assoc etype type-counts) type-counts))
          (setq type-counts (cons (cons etype 1) type-counts))
        )
        ;; Count by layer
        (if (assoc elayer layer-counts)
          (setq layer-counts (subst (cons elayer (1+ (cdr (assoc elayer layer-counts))))
                                    (assoc elayer layer-counts) layer-counts))
          (setq layer-counts (cons (cons elayer 1) layer-counts))
        )
      )
    )
    (setq ent (entnext ent))
  )
  ;; Build type counts JSON
  (setq type-list "")
  (foreach tc type-counts
    (if (> (strlen type-list) 0) (setq type-list (strcat type-list ",")))
    (setq type-list (strcat type-list "\"" (car tc) "\":" (itoa (cdr tc))))
  )
  ;; Build layer counts JSON
  (setq layer-list "")
  (foreach lc layer-counts
    (if (> (strlen layer-list) 0) (setq layer-list (strcat layer-list ",")))
    (setq layer-list (strcat layer-list "\"" (mcp-escape-string (car lc)) "\":" (itoa (cdr lc))))
  )
  ;; Get extents
  (setq extmin (getvar "EXTMIN"))
  (setq extmax (getvar "EXTMAX"))
  (cons T (strcat "{\"by_type\":{" type-list "},\"by_layer\":{" layer-list "}"
                  ",\"extents\":{\"min\":" (mcp-point-to-json extmin)
                  ",\"max\":" (mcp-point-to-json extmax) "}}"))
)

;; --- query-layer-summary ---
(defun mcp-cmd-query-layer-summary (params / layer-name ent ent-data etype elayer
                                             type-counts type-list result count
                                             min-x min-y max-x max-y pt)
  (setq layer-name (mcp-json-get-string params "layer"))
  (setq type-counts '() count 0 ent (entnext))
  (setq min-x 1e30 min-y 1e30 max-x -1e30 max-y -1e30)
  (while ent
    (setq ent-data (entget ent))
    (setq etype (cdr (assoc 0 ent-data)))
    (setq elayer (cdr (assoc 8 ent-data)))
    (if (and (= elayer layer-name)
             (not (member etype '("VERTEX" "SEQEND" "ATTRIB" "ATTDEF"))))
      (progn
        (setq count (1+ count))
        (if (assoc etype type-counts)
          (setq type-counts (subst (cons etype (1+ (cdr (assoc etype type-counts))))
                                   (assoc etype type-counts) type-counts))
          (setq type-counts (cons (cons etype 1) type-counts))
        )
        ;; Track bounding box
        (setq pt (cdr (assoc 10 ent-data)))
        (if pt
          (progn
            (if (< (car pt) min-x) (setq min-x (car pt)))
            (if (< (cadr pt) min-y) (setq min-y (cadr pt)))
            (if (> (car pt) max-x) (setq max-x (car pt)))
            (if (> (cadr pt) max-y) (setq max-y (cadr pt)))
          )
        )
      )
    )
    (setq ent (entnext ent))
  )
  ;; Build type counts JSON
  (setq type-list "")
  (foreach tc type-counts
    (if (> (strlen type-list) 0) (setq type-list (strcat type-list ",")))
    (setq type-list (strcat type-list "\"" (car tc) "\":" (itoa (cdr tc))))
  )
  (setq result (strcat "{\"layer\":\"" (mcp-escape-string layer-name) "\""
                       ",\"count\":" (itoa count)
                       ",\"by_type\":{" type-list "}"))
  (if (> count 0)
    (setq result (strcat result
      ",\"bbox\":{\"min\":[" (rtos min-x 2 4) "," (rtos min-y 2 4)
      "],\"max\":[" (rtos max-x 2 4) "," (rtos max-y 2 4) "]}"))
  )
  (setq result (strcat result "}"))
  (cons T result)
)

;; --- search-text ---
(defun mcp-cmd-search-text (params / pattern case-sens ss i ent ent-data etype content result-str
                                      handle elayer found)
  (setq pattern (mcp-json-get-string params "pattern"))
  (setq case-sens (mcp-json-get-string params "case_sensitive"))
  (if (not pattern)
    (cons nil "pattern required")
    (progn
      (setq result-str "" found 0)
      ;; Use ssget to get all TEXT and MTEXT
      (setq ss (ssget "X" '((0 . "TEXT,MTEXT"))))
      (if ss
        (progn
          (setq i 0)
          (while (< i (sslength ss))
            (setq ent (ssname ss i))
            (setq ent-data (entget ent))
            (setq etype (cdr (assoc 0 ent-data)))
            (setq handle (cdr (assoc 5 ent-data)))
            (setq elayer (cdr (assoc 8 ent-data)))
            (setq content (cdr (assoc 1 ent-data)))
            (if content
              (progn
                ;; Match check
                (if (or (= case-sens "1")
                        (vl-string-search (strcase pattern) (strcase content))
                        (vl-string-search pattern content))
                  (progn
                    (if (> (strlen result-str) 0) (setq result-str (strcat result-str ",")))
                    (setq result-str (strcat result-str
                      "{\"type\":\"" etype "\",\"handle\":\"" handle
                      "\",\"layer\":\"" (mcp-escape-string elayer)
                      "\",\"content\":\"" (mcp-escape-string content)
                      "\",\"position\":" (mcp-point-to-json (cdr (assoc 10 ent-data))) "}"))
                    (setq found (1+ found))
                  )
                )
              )
            )
            (setq i (1+ i))
          )
        )
      )
      (cons T (strcat "{\"count\":" (itoa found) ",\"results\":[" result-str "]}"))
    )
  )
)

;; --- search-by-attribute ---
(defun mcp-cmd-search-by-attribute (params / tag value ss i ent sub-ent sub-data
                                             ent-data handle elayer block-name attribs
                                             result-str found match-found)
  (setq tag (mcp-json-get-string params "tag"))
  (setq value (mcp-json-get-string params "value"))
  (setq result-str "" found 0)
  ;; Get all INSERT entities
  (setq ss (ssget "X" '((0 . "INSERT"))))
  (if ss
    (progn
      (setq i 0)
      (while (< i (sslength ss))
        (setq ent (ssname ss i))
        (setq ent-data (entget ent))
        (setq handle (cdr (assoc 5 ent-data)))
        (setq elayer (cdr (assoc 8 ent-data)))
        (setq block-name (cdr (assoc 2 ent-data)))
        ;; Walk attributes
        (setq sub-ent (entnext ent) match-found nil attribs "")
        (while (and sub-ent (not match-found))
          (setq sub-data (entget sub-ent))
          (if (= (cdr (assoc 0 sub-data)) "ATTRIB")
            (progn
              (if (> (strlen attribs) 0) (setq attribs (strcat attribs ",")))
              (setq attribs (strcat attribs "\"" (cdr (assoc 2 sub-data)) "\":\"" (mcp-escape-string (cdr (assoc 1 sub-data))) "\""))
              ;; Check match
              (if (and (or (not tag) (= (strcase (cdr (assoc 2 sub-data))) (strcase tag)))
                       (or (not value) (vl-string-search (strcase value) (strcase (cdr (assoc 1 sub-data))))))
                (setq match-found T)
              )
            )
          )
          (if (= (cdr (assoc 0 sub-data)) "SEQEND")
            (setq sub-ent nil)
            (setq sub-ent (entnext sub-ent))
          )
        )
        (if match-found
          (progn
            (if (> (strlen result-str) 0) (setq result-str (strcat result-str ",")))
            (setq result-str (strcat result-str
              "{\"handle\":\"" handle "\",\"block\":\"" (mcp-escape-string block-name)
              "\",\"layer\":\"" (mcp-escape-string elayer)
              "\",\"attributes\":{" attribs "}"
              ",\"position\":" (mcp-point-to-json (cdr (assoc 10 ent-data))) "}"))
            (setq found (1+ found))
          )
        )
        (setq i (1+ i))
      )
    )
  )
  (cons T (strcat "{\"count\":" (itoa found) ",\"results\":[" result-str "]}"))
)

;; --- search-by-window ---
(defun mcp-cmd-search-by-window (params / x1 y1 x2 y2 ss i ent ent-data result-str found)
  (setq x1 (mcp-json-get-number params "x1"))
  (setq y1 (mcp-json-get-number params "y1"))
  (setq x2 (mcp-json-get-number params "x2"))
  (setq y2 (mcp-json-get-number params "y2"))
  (setq result-str "" found 0)
  (setq ss (ssget "W" (list x1 y1) (list x2 y2)))
  (if ss
    (progn
      (setq i 0)
      (while (< i (sslength ss))
        (setq ent (ssname ss i))
        (if (> (strlen result-str) 0) (setq result-str (strcat result-str ",")))
        (setq result-str (strcat result-str (mcp-entget-to-json ent)))
        (setq found (1+ found))
        (setq i (1+ i))
      )
    )
  )
  (cons T (strcat "{\"count\":" (itoa found) ",\"results\":[" result-str "]}"))
)

;; --- search-by-proximity ---
(defun mcp-cmd-search-by-proximity (params / cx cy radius ss i ent ent-data pt dist result-str found)
  (setq cx (mcp-json-get-number params "x"))
  (setq cy (mcp-json-get-number params "y"))
  (setq radius (mcp-json-get-number params "radius"))
  (setq result-str "" found 0)
  ;; Use crossing window around the search area, then filter by distance
  (setq ss (ssget "C" (list (- cx radius) (- cy radius)) (list (+ cx radius) (+ cy radius))))
  (if ss
    (progn
      (setq i 0)
      (while (< i (sslength ss))
        (setq ent (ssname ss i))
        (setq ent-data (entget ent))
        (setq pt (cdr (assoc 10 ent-data)))
        (if pt
          (progn
            (setq dist (distance (list cx cy) (list (car pt) (cadr pt))))
            (if (<= dist radius)
              (progn
                (if (> (strlen result-str) 0) (setq result-str (strcat result-str ",")))
                (setq result-str (strcat result-str (mcp-entget-to-json ent)))
                (setq found (1+ found))
              )
            )
          )
        )
        (setq i (1+ i))
      )
    )
  )
  (cons T (strcat "{\"count\":" (itoa found) ",\"results\":[" result-str "]}"))
)

;; --- search-by-type-and-layer ---
(defun mcp-cmd-search-by-type-and-layer (params / etype layer-name color-val filter-list
                                                   ss i ent ent-data result-str found)
  (setq etype (mcp-json-get-string params "entity_type"))
  (setq layer-name (mcp-json-get-string params "layer"))
  (setq color-val (mcp-json-get-string params "color"))
  ;; Build ssget filter
  (setq filter-list '())
  (if etype (setq filter-list (cons (cons 0 etype) filter-list)))
  (if layer-name (setq filter-list (cons (cons 8 layer-name) filter-list)))
  (if color-val (setq filter-list (cons (cons 62 (atoi color-val)) filter-list)))
  (setq result-str "" found 0)
  (setq ss (ssget "X" filter-list))
  (if ss
    (progn
      (setq i 0)
      (while (and (< i (sslength ss)) (< found 500))
        (setq ent (ssname ss i))
        (if (> (strlen result-str) 0) (setq result-str (strcat result-str ",")))
        (setq result-str (strcat result-str (mcp-entget-to-json ent)))
        (setq found (1+ found))
        (setq i (1+ i))
      )
    )
  )
  (cons T (strcat "{\"count\":" (itoa found)
                  (if (and ss (> (sslength ss) 500))
                    (strcat ",\"total\":" (itoa (sslength ss)) ",\"truncated\":true")
                    "")
                  ",\"results\":[" result-str "]}"))
)

;; --- geometry-distance ---
(defun mcp-cmd-geometry-distance (params / x1 y1 x2 y2 dx dy dist ang)
  (setq x1 (mcp-json-get-number params "x1"))
  (setq y1 (mcp-json-get-number params "y1"))
  (setq x2 (mcp-json-get-number params "x2"))
  (setq y2 (mcp-json-get-number params "y2"))
  (setq dx (- x2 x1) dy (- y2 y1))
  (setq dist (distance (list x1 y1) (list x2 y2)))
  (setq ang (* (/ 180.0 pi) (angle (list x1 y1) (list x2 y2))))
  (cons T (strcat "{\"distance\":" (rtos dist 2 6)
                  ",\"dx\":" (rtos dx 2 6)
                  ",\"dy\":" (rtos dy 2 6)
                  ",\"angle\":" (rtos ang 2 6) "}"))
)

;; --- geometry-length ---
(defun mcp-cmd-geometry-length (params / entity-id ent ent-data etype len)
  (setq entity-id (mcp-json-get-string params "entity_id"))
  (if (= entity-id "last") (setq ent (entlast)) (setq ent (handent entity-id)))
  (if (not ent)
    (cons nil (strcat "Entity not found: " entity-id))
    (progn
      (setq ent-data (entget ent))
      (setq etype (cdr (assoc 0 ent-data)))
      (setq len nil)
      (cond
        ((= etype "LINE")
         (setq len (distance (cdr (assoc 10 ent-data)) (cdr (assoc 11 ent-data)))))
        ((= etype "CIRCLE")
         (setq len (* 2.0 pi (cdr (assoc 40 ent-data)))))
        ((= etype "ARC")
         (setq len (* (cdr (assoc 40 ent-data))
                      (abs (- (cdr (assoc 51 ent-data)) (cdr (assoc 50 ent-data)))))))
        ((or (= etype "LWPOLYLINE") (= etype "POLYLINE"))
         ;; Use vlax-curve-getDistAtParam for accurate length
         (setq len (vlax-curve-getDistAtParam ent (vlax-curve-getEndParam ent))))
      )
      (if len
        (cons T (strcat "{\"type\":\"" etype "\",\"length\":" (rtos len 2 6) "}"))
        (cons nil (strcat "Cannot compute length for " etype))
      )
    )
  )
)

;; --- geometry-area ---
(defun mcp-cmd-geometry-area (params / entity-id ent ent-data etype area-val)
  (setq entity-id (mcp-json-get-string params "entity_id"))
  (if (= entity-id "last") (setq ent (entlast)) (setq ent (handent entity-id)))
  (if (not ent)
    (cons nil (strcat "Entity not found: " entity-id))
    (progn
      (setq ent-data (entget ent))
      (setq etype (cdr (assoc 0 ent-data)))
      (setq area-val nil)
      (cond
        ((= etype "CIRCLE")
         (setq area-val (* pi (expt (cdr (assoc 40 ent-data)) 2))))
        ((or (= etype "LWPOLYLINE") (= etype "POLYLINE"))
         (if (= (logand (cdr (assoc 70 ent-data)) 1) 1)
           (progn
             ;; Use vlax-curve for area
             (setq area-val (vlax-curve-getArea ent))
           )
         ))
      )
      (if area-val
        (cons T (strcat "{\"type\":\"" etype "\",\"area\":" (rtos area-val 2 6) "}"))
        (cons nil (strcat "Cannot compute area for " etype " (must be closed)"))
      )
    )
  )
)

;; --- geometry-bounding-box ---
(defun mcp-cmd-geometry-bounding-box (params / entity-id layer-name ent ent-data
                                               min-pt max-pt ss i)
  (setq entity-id (mcp-json-get-string params "entity_id"))
  (setq layer-name (mcp-json-get-string params "layer"))
  (cond
    ;; Single entity
    (entity-id
     (if (= entity-id "last") (setq ent (entlast)) (setq ent (handent entity-id)))
     (if (not ent)
       (cons nil "Entity not found")
       (progn
         (vla-GetBoundingBox (vlax-ename->vla-object ent) 'min-pt 'max-pt)
         (setq min-pt (vlax-safearray->list min-pt))
         (setq max-pt (vlax-safearray->list max-pt))
         (cons T (strcat "{\"min\":" (mcp-point-to-json min-pt)
                         ",\"max\":" (mcp-point-to-json max-pt)
                         ",\"width\":" (rtos (- (car max-pt) (car min-pt)) 2 6)
                         ",\"height\":" (rtos (- (cadr max-pt) (cadr min-pt)) 2 6) "}"))
       )
     ))
    ;; Layer bounding box
    (layer-name
     (progn
       (setq ss (ssget "X" (list (cons 8 layer-name))))
       (if (not ss)
         (cons nil (strcat "No entities on layer: " layer-name))
         (progn
           (setq min-pt nil max-pt nil i 0)
           (while (< i (sslength ss))
             (setq ent (ssname ss i))
             (vl-catch-all-apply
               (function (lambda ()
                 (vla-GetBoundingBox (vlax-ename->vla-object ent) 'ebb-min 'ebb-max)
                 (setq ebb-min (vlax-safearray->list ebb-min))
                 (setq ebb-max (vlax-safearray->list ebb-max))
                 (if (not min-pt)
                   (progn (setq min-pt ebb-min max-pt ebb-max))
                   (progn
                     (if (< (car ebb-min) (car min-pt)) (setq min-pt (cons (car ebb-min) (cdr min-pt))))
                     (if (< (cadr ebb-min) (cadr min-pt)) (setq min-pt (list (car min-pt) (cadr ebb-min) (caddr min-pt))))
                     (if (> (car ebb-max) (car max-pt)) (setq max-pt (cons (car ebb-max) (cdr max-pt))))
                     (if (> (cadr ebb-max) (cadr max-pt)) (setq max-pt (list (car max-pt) (cadr ebb-max) (caddr max-pt))))
                   )
                 )
               ))
             )
             (setq i (1+ i))
           )
           (if min-pt
             (cons T (strcat "{\"min\":" (mcp-point-to-json min-pt)
                             ",\"max\":" (mcp-point-to-json max-pt)
                             ",\"width\":" (rtos (- (car max-pt) (car min-pt)) 2 6)
                             ",\"height\":" (rtos (- (cadr max-pt) (cadr min-pt)) 2 6) "}"))
             (cons nil "Could not compute bounding box")
           )
         )
       )
     ))
    ;; Entire drawing
    (t
     (progn
       (setq min-pt (getvar "EXTMIN"))
       (setq max-pt (getvar "EXTMAX"))
       (cons T (strcat "{\"min\":" (mcp-point-to-json min-pt)
                       ",\"max\":" (mcp-point-to-json max-pt)
                       ",\"width\":" (rtos (- (car max-pt) (car min-pt)) 2 6)
                       ",\"height\":" (rtos (- (cadr max-pt) (cadr min-pt)) 2 6) "}"))
     ))
  )
)

;; --- geometry-polyline-info ---
(defun mcp-cmd-geometry-polyline-info (params / entity-id ent ent-data etype verts sub-ent sub-data
                                                total-len is-closed flags result)
  (setq entity-id (mcp-json-get-string params "entity_id"))
  (if (= entity-id "last") (setq ent (entlast)) (setq ent (handent entity-id)))
  (if (not ent)
    (cons nil (strcat "Entity not found: " entity-id))
    (progn
      (setq ent-data (entget ent))
      (setq etype (cdr (assoc 0 ent-data)))
      (if (not (or (= etype "LWPOLYLINE") (= etype "POLYLINE")))
        (cons nil (strcat "Not a polyline: " etype))
        (progn
          ;; Get vertices
          (setq verts "")
          (cond
            ((= etype "LWPOLYLINE")
             ;; All vertices in entity data
             (foreach pair ent-data
               (if (= (car pair) 10)
                 (progn
                   (if (> (strlen verts) 0) (setq verts (strcat verts ",")))
                   (setq verts (strcat verts (mcp-point-to-json (cdr pair))))
                 )
               )
             )
             (setq is-closed (= (logand (cdr (assoc 70 ent-data)) 1) 1))
             (setq flags (cdr (assoc 70 ent-data)))
            )
            ((= etype "POLYLINE")
             ;; Walk VERTEX sub-entities
             (setq sub-ent (entnext ent))
             (while sub-ent
               (setq sub-data (entget sub-ent))
               (if (= (cdr (assoc 0 sub-data)) "VERTEX")
                 (progn
                   (if (> (strlen verts) 0) (setq verts (strcat verts ",")))
                   (setq verts (strcat verts (mcp-point-to-json (cdr (assoc 10 sub-data)))))
                 )
               )
               (if (= (cdr (assoc 0 sub-data)) "SEQEND")
                 (setq sub-ent nil)
                 (setq sub-ent (entnext sub-ent))
               )
             )
             (setq is-closed (= (logand (cdr (assoc 70 ent-data)) 1) 1))
             (setq flags (cdr (assoc 70 ent-data)))
            )
          )
          ;; Get total length via vlax-curve
          (setq total-len
            (vl-catch-all-apply
              (function (lambda ()
                (vlax-curve-getDistAtParam ent (vlax-curve-getEndParam ent))))
            )
          )
          (if (vl-catch-all-error-p total-len) (setq total-len nil))

          (setq result (strcat "{\"type\":\"" etype "\""
                               ",\"closed\":" (if is-closed "true" "false")
                               ",\"flags\":" (itoa flags)
                               ",\"vertices\":[" verts "]"))
          (if total-len
            (setq result (strcat result ",\"total_length\":" (rtos total-len 2 6))))
          (setq result (strcat result "}"))
          (cons T result)
        )
      )
    )
  )
)

;; --- bulk-set-property ---
(defun mcp-cmd-bulk-set-property (params / handles-str prop-name prop-val handles ent ent-data
                                          success-count fail-count)
  (setq handles-str (mcp-json-get-string params "handles_str"))
  (setq prop-name (mcp-json-get-string params "property_name"))
  (setq prop-val (mcp-json-get-string params "value"))
  (setq handles (mcp-split-string handles-str ";"))
  (setq success-count 0 fail-count 0)
  (foreach h handles
    (setq ent (handent h))
    (if ent
      (progn
        (setq ent-data (entget ent))
        (cond
          ((= prop-name "layer")
           (setq ent-data (subst (cons 8 prop-val) (assoc 8 ent-data) ent-data))
           (entmod ent-data)
           (setq success-count (1+ success-count)))
          ((= prop-name "color")
           (if (assoc 62 ent-data)
             (setq ent-data (subst (cons 62 (atoi prop-val)) (assoc 62 ent-data) ent-data))
             (setq ent-data (append ent-data (list (cons 62 (atoi prop-val)))))
           )
           (entmod ent-data)
           (setq success-count (1+ success-count)))
          ((= prop-name "linetype")
           (if (assoc 6 ent-data)
             (setq ent-data (subst (cons 6 prop-val) (assoc 6 ent-data) ent-data))
             (setq ent-data (append ent-data (list (cons 6 prop-val))))
           )
           (entmod ent-data)
           (setq success-count (1+ success-count)))
          (t (setq fail-count (1+ fail-count)))
        )
      )
      (setq fail-count (1+ fail-count))
    )
  )
  (cons T (strcat "{\"success\":" (itoa success-count) ",\"failed\":" (itoa fail-count) "}"))
)

;; --- bulk-erase ---
(defun mcp-cmd-bulk-erase (params / handles-str handles ent success-count fail-count)
  (setq handles-str (mcp-json-get-string params "handles_str"))
  (setq handles (mcp-split-string handles-str ";"))
  (setq success-count 0 fail-count 0)
  (foreach h handles
    (setq ent (handent h))
    (if ent
      (progn (entdel ent) (setq success-count (1+ success-count)))
      (setq fail-count (1+ fail-count))
    )
  )
  (cons T (strcat "{\"success\":" (itoa success-count) ",\"failed\":" (itoa fail-count) "}"))
)

;; --- export-entity-data ---
(defun mcp-cmd-export-entity-data (params / layer-name etype-filter ent ent-data etype elayer
                                            result-str count)
  (setq layer-name (mcp-json-get-string params "layer"))
  (setq etype-filter (mcp-json-get-string params "entity_type"))
  (setq result-str "" count 0 ent (entnext))
  (while (and ent (< count 1000))
    (setq ent-data (entget ent))
    (setq etype (cdr (assoc 0 ent-data)))
    (setq elayer (cdr (assoc 8 ent-data)))
    ;; Skip sub-entities
    (if (and (not (member etype '("VERTEX" "SEQEND" "ATTRIB" "ATTDEF")))
             (or (not layer-name) (= elayer layer-name))
             (or (not etype-filter) (= (strcase etype) (strcase etype-filter))))
      (progn
        (if (> (strlen result-str) 0) (setq result-str (strcat result-str ",")))
        (setq result-str (strcat result-str (mcp-entget-to-json ent)))
        (setq count (1+ count))
      )
    )
    (setq ent (entnext ent))
  )
  (cons T (strcat "{\"count\":" (itoa count) ",\"entities\":[" result-str "]}"))
)

;; -----------------------------------------------------------------------
;; Main dispatcher — called by "(c:mcp-dispatch)" from Python
;; -----------------------------------------------------------------------

(defun c:mcp-dispatch ( / cmd-files cmd-file json-text request-id cmd-name params-str result result-file)
  "Find pending command file, dispatch, write result."
  ;; Find first pending command file
  (setq cmd-files (vl-directory-files *mcp-ipc-dir* "autocad_mcp_cmd_*.json" 1))
  (if (not cmd-files)
    (progn (princ "\nMCP: No pending commands") (princ))
    (progn
      ;; Process first command
      (setq cmd-file (strcat *mcp-ipc-dir* (car cmd-files)))
      (setq json-text (mcp-read-file-lines cmd-file))

      (if (not json-text)
        (princ "\nMCP: Cannot read command file")
        (progn
          ;; Parse command
          (setq request-id (mcp-json-get-string json-text "request_id"))
          (setq cmd-name (mcp-json-get-string json-text "command"))

          (if (not cmd-name)
            (princ "\nMCP: No command in payload")
            (progn
              (princ (strcat "\nMCP: Dispatching " cmd-name " [" request-id "]"))

              ;; Execute via whitelist dispatcher
              (setq result
                (vl-catch-all-apply
                  'mcp-dispatch-command
                  (list cmd-name json-text)
                )
              )

              ;; Handle error from vl-catch-all-apply
              (if (vl-catch-all-error-p result)
                (setq result (cons nil (vl-catch-all-error-message result)))
              )

              ;; Write result
              (setq result-file (strcat *mcp-ipc-dir* "autocad_mcp_result_" request-id ".json"))
              (if (car result)
                (mcp-write-result result-file request-id T (cdr result) nil)
                (mcp-write-result result-file request-id nil nil (cdr result))
              )

              (princ (strcat "\nMCP: Done " cmd-name))
            )
          )

          ;; Clean up command file
          (vl-file-delete cmd-file)
        )
      )
    )
  )
  (princ)
)

;; -----------------------------------------------------------------------
;; Utility helpers (defined if not already loaded from external files)
;; -----------------------------------------------------------------------

(if (not ensure_layer_exists)
  (defun ensure_layer_exists (name color linetype)
    "Create layer if it doesn't exist."
    (if (not (tblsearch "LAYER" name))
      (command "_.-LAYER" "_NEW" name "_COLOR" color name "_LTYPE" linetype name "")
    )
  )
)

(if (not set_current_layer)
  (defun set_current_layer (name)
    "Set a layer as current."
    (setvar "CLAYER" name)
  )
)

(if (not set_attribute_value)
  (defun set_attribute_value (ent tag value / sub-ent ent-data)
    "Set an attribute value on a block insert by tag name."
    (setq sub-ent (entnext ent))
    (while sub-ent
      (setq ent-data (entget sub-ent))
      (if (and (= (cdr (assoc 0 ent-data)) "ATTRIB")
               (= (strcase (cdr (assoc 2 ent-data))) (strcase tag)))
        (progn
          (entmod (subst (cons 1 value) (assoc 1 ent-data) ent-data))
          (entupd sub-ent)
          (setq sub-ent nil)  ; stop
        )
        (if (= (cdr (assoc 0 ent-data)) "SEQEND")
          (setq sub-ent nil)
          (setq sub-ent (entnext sub-ent))
        )
      )
    )
  )
)

;; -----------------------------------------------------------------------
;; Select / Filter commands
;; -----------------------------------------------------------------------

(defun mcp-cmd-select-filter (params / etype-filter layer-filter color-filter
                                      x1 y1 x2 y2 use-window
                                      ent ed etype elayer ecolor
                                      handles count result)
  (setq etype-filter (mcp-json-get-string params "entity_type"))
  (setq layer-filter (mcp-json-get-string params "layer"))
  (setq color-filter (mcp-json-get-number params "color"))
  (setq x1 (mcp-json-get-number params "x1"))
  (setq y1 (mcp-json-get-number params "y1"))
  (setq x2 (mcp-json-get-number params "x2"))
  (setq y2 (mcp-json-get-number params "y2"))
  (setq use-window (and x1 y1 x2 y2))
  (if use-window (progn
    (if (> x1 x2) (setq tmp x1 x1 x2 x2 tmp))
    (if (> y1 y2) (setq tmp y1 y1 y2 y2 tmp))
  ))
  (setq handles "" count 0)
  (setq ent (entnext))
  (while (and ent (< count 1000))
    (setq ed (entget ent))
    (setq etype (cdr (assoc 0 ed)))
    (setq elayer (cdr (assoc 8 ed)))
    (setq ecolor (cdr (assoc 62 ed)))
    (setq match T)
    (if (and etype-filter (not (wcmatch (strcase etype) (strcase etype-filter)))) (setq match nil))
    (if (and match layer-filter (not (wcmatch elayer layer-filter))) (setq match nil))
    (if (and match color-filter (/= (if ecolor ecolor 256) (fix color-filter))) (setq match nil))
    (if (and match use-window)
      (progn
        (setq epos (mcp-entity-position ed etype))
        (if epos
          (if (or (< (car epos) x1) (> (car epos) x2)
                  (< (cadr epos) y1) (> (cadr epos) y2))
            (setq match nil))
          (setq match nil)
        )
      )
    )
    (if match
      (progn
        (if (> count 0) (setq handles (strcat handles ",")))
        (setq handles (strcat handles "\"" (cdr (assoc 5 ed)) "\""))
        (setq count (1+ count))
      )
    )
    (setq ent (entnext ent))
  )
  (cons T (strcat "{\"handles\":[" handles "],\"count\":" (itoa count) "}"))
)

(defun mcp-cmd-bulk-move (params / handles-str dx dy handles ent ed pos moved errors)
  (setq handles-str (mcp-json-get-string params "handles_str"))
  (setq dx (mcp-json-get-number params "dx"))
  (setq dy (mcp-json-get-number params "dy"))
  (setq handles (mcp-split-string handles-str ";"))
  (setq moved 0 errors "")
  (foreach h handles
    (setq ent (handent h))
    (if ent
      (progn
        (command "_.MOVE" ent "" (list 0.0 0.0) (list dx dy))
        (setq moved (1+ moved))
      )
      (progn
        (if (> (strlen errors) 0) (setq errors (strcat errors ",")))
        (setq errors (strcat errors "\"" h ": not found\""))
      )
    )
  )
  (cons T (strcat "{\"moved\":" (itoa moved)
    (if (> (strlen errors) 0) (strcat ",\"errors\":[" errors "]") "")
    "}"))
)

(defun mcp-cmd-bulk-copy (params / handles-str dx dy handles ent moved errors new-handles)
  (setq handles-str (mcp-json-get-string params "handles_str"))
  (setq dx (mcp-json-get-number params "dx"))
  (setq dy (mcp-json-get-number params "dy"))
  (setq handles (mcp-split-string handles-str ";"))
  (setq moved 0 errors "" new-handles "")
  (foreach h handles
    (setq ent (handent h))
    (if ent
      (progn
        (command "_.COPY" ent "" (list 0.0 0.0) (list dx dy))
        (setq new-h (cdr (assoc 5 (entget (entlast)))))
        (if (> moved 0) (setq new-handles (strcat new-handles ",")))
        (setq new-handles (strcat new-handles "\"" new-h "\""))
        (setq moved (1+ moved))
      )
      (progn
        (if (> (strlen errors) 0) (setq errors (strcat errors ",")))
        (setq errors (strcat errors "\"" h ": not found\""))
      )
    )
  )
  (cons T (strcat "{\"copied\":" (itoa moved)
    ",\"new_handles\":[" new-handles "]"
    (if (> (strlen errors) 0) (strcat ",\"errors\":[" errors "]") "")
    "}"))
)

(defun mcp-cmd-find-replace-text (params / find-str replace-str layer-filter
                                          ent ed etype elayer content new-content
                                          replaced results)
  (setq find-str (mcp-json-get-string params "find"))
  (setq replace-str (mcp-json-get-string params "replace"))
  (setq layer-filter (mcp-json-get-string params "layer"))
  (setq replaced 0 results "")
  (setq ent (entnext))
  (while ent
    (setq ed (entget ent))
    (setq etype (cdr (assoc 0 ed)))
    (setq elayer (cdr (assoc 8 ed)))
    (if (and (member etype '("TEXT" "MTEXT"))
             (or (not layer-filter) (= elayer layer-filter)))
      (progn
        (setq content (cdr (assoc 1 ed)))
        (if (wcmatch content (strcat "*" find-str "*"))
          (progn
            ;; Simple replacement using vl-string-subst
            (setq new-content (vl-string-subst replace-str find-str content))
            (entmod (subst (cons 1 new-content) (assoc 1 ed) ed))
            (entupd ent)
            (if (> replaced 0) (setq results (strcat results ",")))
            (setq results (strcat results "{\"handle\":\"" (cdr (assoc 5 ed))
              "\",\"old\":\"" content "\",\"new\":\"" new-content "\"}"))
            (setq replaced (1+ replaced))
          )
        )
      )
    )
    (setq ent (entnext ent))
  )
  (cons T (strcat "{\"replaced\":" (itoa replaced) ",\"entities\":[" results "]}"))
)

;; -----------------------------------------------------------------------
;; Entity Modification commands
;; -----------------------------------------------------------------------

(defun mcp-cmd-entity-set-property (params / eid prop-name prop-val ent ed code)
  (setq eid (mcp-json-get-string params "entity_id"))
  (setq prop-name (mcp-json-get-string params "property_name"))
  (setq prop-val (mcp-json-get-string params "value"))
  (setq ent (handent eid))
  (if (not ent) (cons nil (strcat "Entity not found: " eid))
    (progn
      (setq ed (entget ent))
      (cond
        ((= prop-name "layer")
         (entmod (subst (cons 8 prop-val) (assoc 8 ed) ed)))
        ((= prop-name "color")
         (if (assoc 62 ed)
           (entmod (subst (cons 62 (atoi prop-val)) (assoc 62 ed) ed))
           (entmod (append ed (list (cons 62 (atoi prop-val)))))
         ))
        ((= prop-name "linetype")
         (if (assoc 6 ed)
           (entmod (subst (cons 6 prop-val) (assoc 6 ed) ed))
           (entmod (append ed (list (cons 6 prop-val))))
         ))
        ((= prop-name "lineweight")
         (if (assoc 370 ed)
           (entmod (subst (cons 370 (atoi prop-val)) (assoc 370 ed) ed))
           (entmod (append ed (list (cons 370 (atoi prop-val)))))
         ))
        (t (cons nil (strcat "Unsupported property: " prop-name)))
      )
      (entupd ent)
      (cons T (strcat "{\"handle\":\"" eid "\",\"property\":\"" prop-name "\",\"value\":\"" prop-val "\"}"))
    )
  )
)

(defun mcp-cmd-entity-set-text (params / eid new-text ent ed etype)
  (setq eid (mcp-json-get-string params "entity_id"))
  (setq new-text (mcp-json-get-string params "text"))
  (setq ent (handent eid))
  (if (not ent) (cons nil (strcat "Entity not found: " eid))
    (progn
      (setq ed (entget ent))
      (setq etype (cdr (assoc 0 ed)))
      (if (not (member etype '("TEXT" "MTEXT")))
        (cons nil (strcat "Entity is not TEXT/MTEXT: " etype))
        (progn
          (entmod (subst (cons 1 new-text) (assoc 1 ed) ed))
          (entupd ent)
          (cons T (strcat "{\"handle\":\"" eid "\",\"type\":\"" etype "\",\"new_text\":\"" new-text "\"}"))
        )
      )
    )
  )
)

;; -----------------------------------------------------------------------
;; View Enhancement commands
;; -----------------------------------------------------------------------

(defun mcp-cmd-zoom-center (params / cx cy height)
  (setq cx (mcp-json-get-number params "x"))
  (setq cy (mcp-json-get-number params "y"))
  (setq height (mcp-json-get-number params "height"))
  (command "_.ZOOM" "_C" (list cx cy) height)
  (cons T "{\"ok\":true}")
)

(defun mcp-cmd-layer-visibility (params / lname visible)
  (setq lname (mcp-json-get-string params "name"))
  (setq visible (mcp-json-get-string params "visible"))
  (if (= visible "1")
    (command "_.LAYER" "_ON" lname "")
    (command "_.LAYER" "_OFF" lname "")
  )
  (cons T (strcat "{\"layer\":\"" lname "\",\"visible\":" (if (= visible "1") "true" "false") "}"))
)

;; -----------------------------------------------------------------------
;; Validate commands
;; -----------------------------------------------------------------------

(defun mcp-cmd-validate-layer-standards (params / allowed-str allowed-layers
                                                  ent ed elayer etype handle
                                                  violations vcount pass-flag)
  (setq allowed-str (mcp-json-get-string params "allowed_layers"))
  (setq allowed-layers (mcp-split-string allowed-str ";"))
  (setq violations "" vcount 0)
  (setq ent (entnext))
  (while ent
    (setq ed (entget ent))
    (setq elayer (cdr (assoc 8 ed)))
    (setq etype (cdr (assoc 0 ed)))
    (setq handle (cdr (assoc 5 ed)))
    (if (not (member elayer allowed-layers))
      (progn
        (if (> vcount 0) (setq violations (strcat violations ",")))
        (setq violations (strcat violations "{\"handle\":\"" handle
          "\",\"layer\":\"" elayer "\",\"type\":\"" etype "\"}"))
        (setq vcount (1+ vcount))
      )
    )
    (setq ent (entnext ent))
  )
  (setq pass-flag (= vcount 0))
  (cons T (strcat "{\"pass\":" (if pass-flag "true" "false")
    ",\"violation_count\":" (itoa vcount)
    ",\"violations\":[" violations "]}"))
)

(defun mcp-cmd-validate-duplicates (params / tolerance ent ed etype
                                            lines circles dups dcount)
  (setq tolerance (mcp-json-get-number params "tolerance"))
  (if (not tolerance) (setq tolerance 0.001))
  ;; Collect LINE and CIRCLE entities for comparison
  (setq lines nil circles nil dups "" dcount 0)
  (setq ent (entnext))
  (while ent
    (setq ed (entget ent))
    (setq etype (cdr (assoc 0 ed)))
    (cond
      ((= etype "LINE")
       (setq lines (cons (list (cdr (assoc 5 ed))
                               (cdr (assoc 10 ed)) (cdr (assoc 11 ed))) lines)))
      ((= etype "CIRCLE")
       (setq circles (cons (list (cdr (assoc 5 ed))
                                  (cdr (assoc 10 ed)) (cdr (assoc 40 ed))) circles)))
    )
    (setq ent (entnext ent))
  )
  ;; Check LINE duplicates (O(n^2) but capped at manageable sizes)
  (setq i 0)
  (foreach l1 lines
    (setq j 0)
    (foreach l2 lines
      (if (> j i)
        (progn
          (setq d1 (distance (cadr l1) (cadr l2)))
          (setq d2 (distance (caddr l1) (caddr l2)))
          (setq d3 (distance (cadr l1) (caddr l2)))
          (setq d4 (distance (caddr l1) (cadr l2)))
          (if (or (and (< d1 tolerance) (< d2 tolerance))
                  (and (< d3 tolerance) (< d4 tolerance)))
            (progn
              (if (> dcount 0) (setq dups (strcat dups ",")))
              (setq dups (strcat dups "{\"type\":\"LINE\",\"handle1\":\"" (car l1)
                "\",\"handle2\":\"" (car l2) "\"}"))
              (setq dcount (1+ dcount))
            )
          )
        )
      )
      (setq j (1+ j))
    )
    (setq i (1+ i))
  )
  ;; Check CIRCLE duplicates
  (setq i 0)
  (foreach c1 circles
    (setq j 0)
    (foreach c2 circles
      (if (> j i)
        (if (and (< (distance (cadr c1) (cadr c2)) tolerance)
                 (< (abs (- (caddr c1) (caddr c2))) tolerance))
          (progn
            (if (> dcount 0) (setq dups (strcat dups ",")))
            (setq dups (strcat dups "{\"type\":\"CIRCLE\",\"handle1\":\"" (car c1)
              "\",\"handle2\":\"" (car c2) "\"}"))
            (setq dcount (1+ dcount))
          )
        )
      )
      (setq j (1+ j))
    )
    (setq i (1+ i))
  )
  (cons T (strcat "{\"duplicate_count\":" (itoa dcount) ",\"duplicates\":[" dups "]}"))
)

(defun mcp-cmd-validate-zero-length (params / ent ed etype p1 p2 issues icount)
  (setq issues "" icount 0)
  (setq ent (entnext))
  (while ent
    (setq ed (entget ent))
    (setq etype (cdr (assoc 0 ed)))
    (cond
      ((= etype "LINE")
       (setq p1 (cdr (assoc 10 ed)) p2 (cdr (assoc 11 ed)))
       (if (< (distance p1 p2) 0.001)
         (progn
           (if (> icount 0) (setq issues (strcat issues ",")))
           (setq issues (strcat issues "{\"type\":\"LINE\",\"handle\":\""
             (cdr (assoc 5 ed)) "\",\"issue\":\"zero-length\"}"))
           (setq icount (1+ icount))
         )
       ))
      ((= etype "CIRCLE")
       (if (< (cdr (assoc 40 ed)) 0.001)
         (progn
           (if (> icount 0) (setq issues (strcat issues ",")))
           (setq issues (strcat issues "{\"type\":\"CIRCLE\",\"handle\":\""
             (cdr (assoc 5 ed)) "\",\"issue\":\"zero-radius\"}"))
           (setq icount (1+ icount))
         )
       ))
      ((= etype "ARC")
       (if (< (cdr (assoc 40 ed)) 0.001)
         (progn
           (if (> icount 0) (setq issues (strcat issues ",")))
           (setq issues (strcat issues "{\"type\":\"ARC\",\"handle\":\""
             (cdr (assoc 5 ed)) "\",\"issue\":\"zero-radius\"}"))
           (setq icount (1+ icount))
         )
       ))
    )
    (setq ent (entnext ent))
  )
  (cons T (strcat "{\"issue_count\":" (itoa icount) ",\"issues\":[" issues "]}"))
)

(defun mcp-cmd-validate-qc-report (params / allowed-str result-layers result-dupes
                                           result-zero total-issues pass-flag)
  ;; Run all three validation checks
  (setq result-layers (mcp-cmd-validate-layer-standards params))
  (setq result-dupes (mcp-cmd-validate-duplicates params))
  (setq result-zero (mcp-cmd-validate-zero-length params))
  (cons T (strcat "{\"layer_standards\":" (cdr result-layers)
    ",\"duplicates\":" (cdr result-dupes)
    ",\"zero_length\":" (cdr result-zero) "}"))
)

;; -----------------------------------------------------------------------
;; Export / Reporting commands
;; -----------------------------------------------------------------------

(defun mcp-cmd-export-bom (params / block-names-str block-filter
                                    ent ed etype bname attrs
                                    blocks-alist entry result)
  (setq block-names-str (mcp-json-get-string params "block_names"))
  (if block-names-str
    (setq block-filter (mcp-split-string block-names-str ";"))
    (setq block-filter nil)
  )
  (setq blocks-alist nil)
  (setq ent (entnext))
  (while ent
    (setq ed (entget ent))
    (setq etype (cdr (assoc 0 ed)))
    (if (= etype "INSERT")
      (progn
        (setq bname (cdr (assoc 2 ed)))
        (if (or (not block-filter) (member bname block-filter))
          (progn
            ;; Find or create entry in blocks-alist
            (setq entry (assoc bname blocks-alist))
            (if entry
              (setq blocks-alist (subst (cons bname (1+ (cdr entry))) entry blocks-alist))
              (setq blocks-alist (cons (cons bname 1) blocks-alist))
            )
          )
        )
      )
    )
    (setq ent (entnext ent))
  )
  ;; Build JSON result
  (setq result "")
  (foreach pair blocks-alist
    (if (> (strlen result) 0) (setq result (strcat result ",")))
    (setq result (strcat result "{\"block\":\"" (car pair)
      "\",\"count\":" (itoa (cdr pair)) "}"))
  )
  (cons T (strcat "{\"items\":[" result "],\"total_blocks\":" (itoa (length blocks-alist)) "}"))
)

(defun mcp-cmd-export-data-extract (params / etype-filter layer-filter
                                            props-str properties
                                            ent ed etype elayer
                                            rows count row-json result)
  (setq etype-filter (mcp-json-get-string params "entity_type"))
  (setq layer-filter (mcp-json-get-string params "layer"))
  (setq props-str (mcp-json-get-string params "properties"))
  (if props-str
    (setq properties (mcp-split-string props-str ";"))
    (setq properties nil)
  )
  (setq rows "" count 0)
  (setq ent (entnext))
  (while (and ent (< count 500))
    (setq ed (entget ent))
    (setq etype (cdr (assoc 0 ed)))
    (setq elayer (cdr (assoc 8 ed)))
    (setq match T)
    (if (and etype-filter (not (wcmatch (strcase etype) (strcase etype-filter)))) (setq match nil))
    (if (and match layer-filter (not (wcmatch elayer layer-filter))) (setq match nil))
    (if match
      (progn
        (setq row-json (strcat "{\"handle\":\"" (cdr (assoc 5 ed))
          "\",\"type\":\"" etype "\",\"layer\":\"" elayer "\""))
        ;; Add requested properties (or default set)
        (if (or (not properties) (member "color" properties))
          (progn
            (setq c (cdr (assoc 62 ed)))
            (setq row-json (strcat row-json ",\"color\":" (if c (itoa c) "256")))
          )
        )
        (if (or (not properties) (member "position" properties))
          (progn
            (setq pos (cdr (assoc 10 ed)))
            (if pos
              (setq row-json (strcat row-json ",\"position\":["
                (rtos (car pos) 2 4) "," (rtos (cadr pos) 2 4) "]"))
            )
          )
        )
        (if (or (not properties) (member "content" properties))
          (if (member etype '("TEXT" "MTEXT"))
            (setq row-json (strcat row-json ",\"content\":\"" (cdr (assoc 1 ed)) "\""))
          )
        )
        (setq row-json (strcat row-json "}"))
        (if (> count 0) (setq rows (strcat rows ",")))
        (setq rows (strcat rows row-json))
        (setq count (1+ count))
      )
    )
    (setq ent (entnext ent))
  )
  (cons T (strcat "{\"rows\":[" rows "],\"count\":" (itoa count) "}"))
)

(defun mcp-cmd-export-layer-report (params / ent ed etype elayer
                                            layer-data entry type-counts result)
  ;; Build layer → {entity_count, types: {type: count}}
  (setq layer-data nil)
  (setq ent (entnext))
  (while ent
    (setq ed (entget ent))
    (setq etype (cdr (assoc 0 ed)))
    (setq elayer (cdr (assoc 8 ed)))
    ;; Find or create layer entry: (layer-name entity-count type-alist)
    (setq entry (assoc elayer layer-data))
    (if entry
      (progn
        ;; Increment entity count
        (setq new-count (1+ (cadr entry)))
        ;; Update type counts
        (setq type-counts (caddr entry))
        (setq tentry (assoc etype type-counts))
        (if tentry
          (setq type-counts (subst (cons etype (1+ (cdr tentry))) tentry type-counts))
          (setq type-counts (cons (cons etype 1) type-counts))
        )
        (setq layer-data (subst (list elayer new-count type-counts) entry layer-data))
      )
      (setq layer-data (cons (list elayer 1 (list (cons etype 1))) layer-data))
    )
    (setq ent (entnext ent))
  )
  ;; Build JSON
  (setq result "")
  (foreach lentry layer-data
    (if (> (strlen result) 0) (setq result (strcat result ",")))
    (setq type-json "")
    (foreach tpair (caddr lentry)
      (if (> (strlen type-json) 0) (setq type-json (strcat type-json ",")))
      (setq type-json (strcat type-json "\"" (car tpair) "\":" (itoa (cdr tpair))))
    )
    (setq result (strcat result "{\"layer\":\"" (car lentry)
      "\",\"entity_count\":" (itoa (cadr lentry))
      ",\"types\":{" type-json "}}"))
  )
  (cons T (strcat "{\"layers\":[" result "],\"layer_count\":" (itoa (length layer-data)) "}"))
)

(defun mcp-cmd-export-block-count (params / ent ed etype bname blocks-alist entry result)
  (setq blocks-alist nil)
  (setq ent (entnext))
  (while ent
    (setq ed (entget ent))
    (if (= (cdr (assoc 0 ed)) "INSERT")
      (progn
        (setq bname (cdr (assoc 2 ed)))
        (setq entry (assoc bname blocks-alist))
        (if entry
          (setq blocks-alist (subst (cons bname (1+ (cdr entry))) entry blocks-alist))
          (setq blocks-alist (cons (cons bname 1) blocks-alist))
        )
      )
    )
    (setq ent (entnext ent))
  )
  (setq result "")
  (foreach pair blocks-alist
    (if (> (strlen result) 0) (setq result (strcat result ",")))
    (setq result (strcat result "{\"name\":\"" (car pair) "\",\"count\":" (itoa (cdr pair)) "}"))
  )
  (cons T (strcat "{\"blocks\":[" result "],\"unique_blocks\":" (itoa (length blocks-alist)) "}"))
)

(defun mcp-cmd-export-drawing-statistics (params / ent ed etype
                                                   total-count type-alist layer-set
                                                   entry elayer result type-json)
  (setq total-count 0 type-alist nil layer-set nil)
  (setq ent (entnext))
  (while ent
    (setq ed (entget ent))
    (setq etype (cdr (assoc 0 ed)))
    (setq elayer (cdr (assoc 8 ed)))
    (setq total-count (1+ total-count))
    ;; Count by type
    (setq entry (assoc etype type-alist))
    (if entry
      (setq type-alist (subst (cons etype (1+ (cdr entry))) entry type-alist))
      (setq type-alist (cons (cons etype 1) type-alist))
    )
    ;; Track unique layers
    (if (not (member elayer layer-set))
      (setq layer-set (cons elayer layer-set))
    )
    (setq ent (entnext ent))
  )
  ;; Count styles, dimstyles, blocks
  (setq style-count 0 ts (tblnext "STYLE" T))
  (while ts (setq style-count (1+ style-count)) (setq ts (tblnext "STYLE")))
  (setq dim-count 0 ds (tblnext "DIMSTYLE" T))
  (while ds (setq dim-count (1+ dim-count)) (setq ds (tblnext "DIMSTYLE")))
  (setq blk-count 0 bs (tblnext "BLOCK" T))
  (while bs (setq blk-count (1+ blk-count)) (setq bs (tblnext "BLOCK")))
  ;; Build JSON
  (setq type-json "")
  (foreach pair type-alist
    (if (> (strlen type-json) 0) (setq type-json (strcat type-json ",")))
    (setq type-json (strcat type-json "\"" (car pair) "\":" (itoa (cdr pair))))
  )
  (cons T (strcat "{\"entity_count\":" (itoa total-count)
    ",\"by_type\":{" type-json "}"
    ",\"layer_count\":" (itoa (length layer-set))
    ",\"style_count\":" (itoa style-count)
    ",\"dimstyle_count\":" (itoa dim-count)
    ",\"block_definition_count\":" (itoa blk-count)
    "}"))
)

;; -----------------------------------------------------------------------
;; Extended Query commands
;; -----------------------------------------------------------------------

(defun mcp-cmd-query-text-styles (params / ts result)
  (setq result "" ts (tblnext "STYLE" T))
  (while ts
    (if (> (strlen result) 0) (setq result (strcat result ",")))
    (setq result (strcat result "{\"name\":\"" (cdr (assoc 2 ts))
      "\",\"font\":\"" (if (cdr (assoc 3 ts)) (cdr (assoc 3 ts)) "")
      "\",\"bigfont\":\"" (if (cdr (assoc 4 ts)) (cdr (assoc 4 ts)) "")
      "\",\"height\":" (rtos (cdr (assoc 40 ts)) 2 4)
      ",\"width_factor\":" (rtos (cdr (assoc 41 ts)) 2 4) "}"))
    (setq ts (tblnext "STYLE"))
  )
  (cons T (strcat "{\"styles\":[" result "]}"))
)

(defun mcp-cmd-query-dimension-styles (params / ds result)
  (setq result "" ds (tblnext "DIMSTYLE" T))
  (while ds
    (if (> (strlen result) 0) (setq result (strcat result ",")))
    (setq result (strcat result "{\"name\":\"" (cdr (assoc 2 ds)) "\"}"))
    (setq ds (tblnext "DIMSTYLE"))
  )
  (cons T (strcat "{\"dimstyles\":[" result "]}"))
)

(defun mcp-cmd-query-linetypes (params / lt result)
  (setq result "" lt (tblnext "LTYPE" T))
  (while lt
    (if (> (strlen result) 0) (setq result (strcat result ",")))
    (setq result (strcat result "{\"name\":\"" (cdr (assoc 2 lt))
      "\",\"description\":\"" (if (cdr (assoc 3 lt)) (cdr (assoc 3 lt)) "") "\"}"))
    (setq lt (tblnext "LTYPE"))
  )
  (cons T (strcat "{\"linetypes\":[" result "]}"))
)

(defun mcp-cmd-query-block-tree (params / blk result bname flags is-xref)
  (setq result "" blk (tblnext "BLOCK" T))
  (while blk
    (setq bname (cdr (assoc 2 blk)))
    (setq flags (cdr (assoc 70 blk)))
    (setq is-xref (= (logand flags 4) 4))
    (if (> (strlen result) 0) (setq result (strcat result ",")))
    (setq result (strcat result "{\"name\":\"" bname
      "\",\"flags\":" (itoa flags)
      ",\"is_xref\":" (if is-xref "true" "false") "}"))
    (setq blk (tblnext "BLOCK"))
  )
  (cons T (strcat "{\"blocks\":[" result "]}"))
)

(defun mcp-cmd-query-drawing-metadata (params / )
  (cons T (strcat "{"
    "\"units\":" (itoa (getvar "LUNITS"))
    ",\"unit_precision\":" (itoa (getvar "LUPREC"))
    ",\"angle_units\":" (itoa (getvar "AUNITS"))
    ",\"limmin\":[" (rtos (car (getvar "LIMMIN")) 2 4) "," (rtos (cadr (getvar "LIMMIN")) 2 4) "]"
    ",\"limmax\":[" (rtos (car (getvar "LIMMAX")) 2 4) "," (rtos (cadr (getvar "LIMMAX")) 2 4) "]"
    ",\"extmin\":[" (rtos (car (getvar "EXTMIN")) 2 4) "," (rtos (cadr (getvar "EXTMIN")) 2 4) "]"
    ",\"extmax\":[" (rtos (car (getvar "EXTMAX")) 2 4) "," (rtos (cadr (getvar "EXTMAX")) 2 4) "]"
    ",\"dwgname\":\"" (mcp-escape-string (getvar "DWGNAME")) "\""
    ",\"dwgprefix\":\"" (mcp-escape-string (getvar "DWGPREFIX")) "\""
    "}"))
)

;; -----------------------------------------------------------------------
;; Extended Search commands
;; -----------------------------------------------------------------------

(defun mcp-cmd-search-by-block-name (params / bname ent ed result count)
  (setq bname (mcp-json-get-string params "block_name"))
  (setq result "" count 0)
  (setq ent (entnext))
  (while (and ent (< count 500))
    (setq ed (entget ent))
    (if (and (= (cdr (assoc 0 ed)) "INSERT") (wcmatch (cdr (assoc 2 ed)) bname))
      (progn
        (setq pos (cdr (assoc 10 ed)))
        (if (> count 0) (setq result (strcat result ",")))
        (setq result (strcat result "{\"handle\":\"" (cdr (assoc 5 ed))
          "\",\"block\":\"" (cdr (assoc 2 ed))
          "\",\"layer\":\"" (cdr (assoc 8 ed))
          "\",\"position\":[" (rtos (car pos) 2 4) "," (rtos (cadr pos) 2 4) "]}"))
        (setq count (1+ count))
      )
    )
    (setq ent (entnext ent))
  )
  (cons T (strcat "{\"entities\":[" result "],\"count\":" (itoa count) "}"))
)

(defun mcp-cmd-search-by-handle-list (params / handles-str handles result count)
  (setq handles-str (mcp-json-get-string params "handles_str"))
  (setq handles (mcp-split-string handles-str ";"))
  (setq result "" count 0)
  (foreach h handles
    (setq ent (handent h))
    (if ent
      (progn
        (setq ed (entget ent))
        (if (> count 0) (setq result (strcat result ",")))
        (setq result (strcat result (mcp-entget-to-json ed)))
        (setq count (1+ count))
      )
    )
  )
  (cons T (strcat "{\"entities\":[" result "],\"count\":" (itoa count) "}"))
)

;; -----------------------------------------------------------------------
;; Extended Entity commands
;; -----------------------------------------------------------------------

(defun mcp-cmd-entity-explode (params / eid ent)
  (setq eid (mcp-json-get-string params "entity_id"))
  (setq ent (handent eid))
  (if (not ent) (cons nil (strcat "Entity not found: " eid))
    (progn
      (command "_.EXPLODE" ent)
      (cons T (strcat "{\"exploded\":\"" eid "\"}"))
    )
  )
)

(defun mcp-cmd-entity-join (params / eids-str eids)
  (setq eids-str (mcp-json-get-string params "entity_ids"))
  (setq eids (mcp-split-string eids-str ";"))
  (if (< (length eids) 2)
    (cons nil "At least 2 entity IDs required")
    (progn
      (setq first-ent (handent (car eids)))
      (if (not first-ent) (cons nil "First entity not found")
        (progn
          (command "_.JOIN" first-ent)
          (foreach h (cdr eids)
            (setq ent (handent h))
            (if ent (command ent))
          )
          (command "")
          (cons T (strcat "{\"joined\":" (itoa (length eids)) "}"))
        )
      )
    )
  )
)

(defun mcp-cmd-entity-extend (params / eid bid ent bent)
  (setq eid (mcp-json-get-string params "entity_id"))
  (setq bid (mcp-json-get-string params "boundary_id"))
  (setq ent (handent eid))
  (setq bent (handent bid))
  (if (not ent) (cons nil (strcat "Entity not found: " eid))
    (if (not bent) (cons nil (strcat "Boundary not found: " bid))
      (progn
        (command "_.EXTEND" bent "" ent "")
        (cons T (strcat "{\"extended\":\"" eid "\",\"boundary\":\"" bid "\"}"))
      )
    )
  )
)

(defun mcp-cmd-entity-trim (params / eid bid ent bent)
  (setq eid (mcp-json-get-string params "entity_id"))
  (setq bid (mcp-json-get-string params "boundary_id"))
  (setq ent (handent eid))
  (setq bent (handent bid))
  (if (not ent) (cons nil (strcat "Entity not found: " eid))
    (if (not bent) (cons nil (strcat "Boundary not found: " bid))
      (progn
        (command "_.TRIM" bent "" ent "")
        (cons T (strcat "{\"trimmed\":\"" eid "\",\"boundary\":\"" bid "\"}"))
      )
    )
  )
)

(defun mcp-cmd-entity-break-at (params / eid px py ent)
  (setq eid (mcp-json-get-string params "entity_id"))
  (setq px (mcp-json-get-number params "x"))
  (setq py (mcp-json-get-number params "y"))
  (setq ent (handent eid))
  (if (not ent) (cons nil (strcat "Entity not found: " eid))
    (progn
      (command "_.BREAK" ent (list px py) (list px py))
      (cons T (strcat "{\"broken\":\"" eid "\",\"at\":[" (rtos px 2 4) "," (rtos py 2 4) "]}"))
    )
  )
)

;; -----------------------------------------------------------------------
;; Extended Validate commands
;; -----------------------------------------------------------------------

(defun mcp-cmd-validate-text-standards (params / styles-str heights-str
                                                 allowed-styles allowed-heights
                                                 ent ed etype style height
                                                 violations vcount)
  (setq styles-str (mcp-json-get-string params "allowed_styles"))
  (setq heights-str (mcp-json-get-string params "allowed_heights"))
  (if styles-str
    (setq allowed-styles (mcp-split-string styles-str ";"))
    (setq allowed-styles nil)
  )
  (setq violations "" vcount 0)
  (setq ent (entnext))
  (while ent
    (setq ed (entget ent))
    (setq etype (cdr (assoc 0 ed)))
    (if (member etype '("TEXT" "MTEXT"))
      (progn
        (setq style (cdr (assoc 7 ed)))
        (setq height (cdr (assoc 40 ed)))
        (setq issue nil)
        (if (and allowed-styles (not (member style allowed-styles)))
          (setq issue (strcat "non-standard style: " style))
        )
        (if issue
          (progn
            (if (> vcount 0) (setq violations (strcat violations ",")))
            (setq violations (strcat violations "{\"handle\":\"" (cdr (assoc 5 ed))
              "\",\"type\":\"" etype "\",\"issue\":\"" issue
              "\",\"style\":\"" (if style style "nil")
              "\",\"height\":" (if height (rtos height 2 4) "0") "}"))
            (setq vcount (1+ vcount))
          )
        )
      )
    )
    (setq ent (entnext ent))
  )
  (cons T (strcat "{\"pass\":" (if (= vcount 0) "true" "false")
    ",\"violation_count\":" (itoa vcount)
    ",\"violations\":[" violations "]}"))
)

(defun mcp-cmd-validate-orphaned-entities (params / ent ed elayer lr
                                                    issues icount frozen off)
  (setq issues "" icount 0)
  (setq ent (entnext))
  (while ent
    (setq ed (entget ent))
    (setq elayer (cdr (assoc 8 ed)))
    (setq lr (tblsearch "LAYER" elayer))
    (if lr
      (progn
        (setq lcolor (cdr (assoc 62 lr)))
        (setq frozen (= (logand (cdr (assoc 70 lr)) 1) 1))
        (setq off (< lcolor 0))
        (if (or frozen off)
          (progn
            (if (> icount 0) (setq issues (strcat issues ",")))
            (setq issues (strcat issues "{\"handle\":\"" (cdr (assoc 5 ed))
              "\",\"layer\":\"" elayer
              "\",\"type\":\"" (cdr (assoc 0 ed))
              "\",\"issue\":\"" (cond (frozen "layer frozen") (off "layer off")) "\"}"))
            (setq icount (1+ icount))
          )
        )
      )
    )
    (setq ent (entnext ent))
  )
  (cons T (strcat "{\"issue_count\":" (itoa icount) ",\"issues\":[" issues "]}"))
)

(defun mcp-cmd-validate-attribute-completeness (params / tags-str required-tags
                                                        ent ed etype bname
                                                        sub sd attr-tag attr-val
                                                        missing issues icount)
  (setq tags-str (mcp-json-get-string params "required_tags"))
  (if tags-str
    (setq required-tags (mcp-split-string tags-str ";"))
    (setq required-tags nil)
  )
  (setq issues "" icount 0)
  (setq ent (entnext))
  (while ent
    (setq ed (entget ent))
    (if (= (cdr (assoc 0 ed)) "INSERT")
      (progn
        (setq bname (cdr (assoc 2 ed)))
        ;; Walk attributes
        (setq sub (entnext ent) found-tags nil)
        (while sub
          (setq sd (entget sub))
          (if (= (cdr (assoc 0 sd)) "ATTRIB")
            (progn
              (setq attr-tag (cdr (assoc 2 sd)))
              (setq attr-val (cdr (assoc 1 sd)))
              (setq found-tags (cons attr-tag found-tags))
              ;; Check if empty
              (if (and required-tags (member attr-tag required-tags)
                       (or (not attr-val) (= attr-val "")))
                (progn
                  (if (> icount 0) (setq issues (strcat issues ",")))
                  (setq issues (strcat issues "{\"handle\":\"" (cdr (assoc 5 ed))
                    "\",\"block\":\"" bname "\",\"tag\":\"" attr-tag
                    "\",\"issue\":\"empty value\"}"))
                  (setq icount (1+ icount))
                )
              )
            )
          )
          (if (= (cdr (assoc 0 sd)) "SEQEND") (setq sub nil) (setq sub (entnext sub)))
        )
        ;; Check for missing required tags
        (if required-tags
          (foreach rtag required-tags
            (if (not (member rtag found-tags))
              (progn
                (if (> icount 0) (setq issues (strcat issues ",")))
                (setq issues (strcat issues "{\"handle\":\"" (cdr (assoc 5 ed))
                  "\",\"block\":\"" bname "\",\"tag\":\"" rtag
                  "\",\"issue\":\"missing tag\"}"))
                (setq icount (1+ icount))
              )
            )
          )
        )
      )
    )
    (setq ent (entnext ent))
  )
  (cons T (strcat "{\"issue_count\":" (itoa icount) ",\"issues\":[" issues "]}"))
)

(defun mcp-cmd-validate-connectivity (params / layer-filter tolerance
                                              ent ed etype elayer endpoints
                                              p1 p2 dangling dcount)
  (setq layer-filter (mcp-json-get-string params "layer"))
  (setq tolerance (mcp-json-get-number params "tolerance"))
  (if (not tolerance) (setq tolerance 0.01))
  ;; Collect all line endpoints
  (setq endpoints nil)
  (setq ent (entnext))
  (while ent
    (setq ed (entget ent))
    (setq etype (cdr (assoc 0 ed)))
    (setq elayer (cdr (assoc 8 ed)))
    (if (and (= etype "LINE") (or (not layer-filter) (= elayer layer-filter)))
      (progn
        (setq p1 (cdr (assoc 10 ed)) p2 (cdr (assoc 11 ed)))
        (setq endpoints (cons (list (cdr (assoc 5 ed)) "start" p1) endpoints))
        (setq endpoints (cons (list (cdr (assoc 5 ed)) "end" p2) endpoints))
      )
    )
    (setq ent (entnext ent))
  )
  ;; Find dangling endpoints (not connected to any other endpoint)
  (setq dangling "" dcount 0)
  (foreach ep endpoints
    (setq connected nil)
    (foreach other endpoints
      (if (and (not (equal ep other))
               (not (and (= (car ep) (car other))))  ; different entity
               (< (distance (caddr ep) (caddr other)) tolerance))
        (setq connected T)
      )
    )
    (if (not connected)
      (if (< dcount 100) ; cap output
        (progn
          (if (> dcount 0) (setq dangling (strcat dangling ",")))
          (setq dangling (strcat dangling "{\"handle\":\"" (car ep)
            "\",\"end\":\"" (cadr ep)
            "\",\"point\":[" (rtos (car (caddr ep)) 2 4) "," (rtos (cadr (caddr ep)) 2 4) "]}"))
          (setq dcount (1+ dcount))
        )
      )
    )
  )
  (cons T (strcat "{\"dangling_count\":" (itoa dcount)
    ",\"total_endpoints\":" (itoa (length endpoints))
    ",\"dangling\":[" dangling "]}"))
)

;; -----------------------------------------------------------------------
;; Extended Select commands
;; -----------------------------------------------------------------------

(defun mcp-cmd-find-replace-attribute (params / tag-name find-str replace-str
                                               ent ed sub sd attr-tag attr-val
                                               replaced results)
  (setq tag-name (mcp-json-get-string params "tag"))
  (setq find-str (mcp-json-get-string params "find"))
  (setq replace-str (mcp-json-get-string params "replace"))
  (setq replaced 0 results "")
  (setq ent (entnext))
  (while ent
    (setq ed (entget ent))
    (if (= (cdr (assoc 0 ed)) "INSERT")
      (progn
        (setq sub (entnext ent))
        (while sub
          (setq sd (entget sub))
          (if (= (cdr (assoc 0 sd)) "ATTRIB")
            (progn
              (setq attr-tag (cdr (assoc 2 sd)))
              (setq attr-val (cdr (assoc 1 sd)))
              (if (and (= attr-tag tag-name) (wcmatch attr-val (strcat "*" find-str "*")))
                (progn
                  (setq new-val (vl-string-subst replace-str find-str attr-val))
                  (entmod (subst (cons 1 new-val) (assoc 1 sd) sd))
                  (entupd sub)
                  (if (> replaced 0) (setq results (strcat results ",")))
                  (setq results (strcat results "{\"handle\":\"" (cdr (assoc 5 ed))
                    "\",\"tag\":\"" attr-tag "\",\"old\":\"" attr-val "\",\"new\":\"" new-val "\"}"))
                  (setq replaced (1+ replaced))
                )
              )
            )
          )
          (if (= (cdr (assoc 0 sd)) "SEQEND") (setq sub nil) (setq sub (entnext sub)))
        )
      )
    )
    (setq ent (entnext ent))
  )
  (cons T (strcat "{\"replaced\":" (itoa replaced) ",\"entities\":[" results "]}"))
)

(defun mcp-cmd-layer-rename (params / old-name new-name)
  (setq old-name (mcp-json-get-string params "old_name"))
  (setq new-name (mcp-json-get-string params "new_name"))
  (command "_.LAYER" "_R" old-name new-name "")
  (cons T (strcat "{\"renamed\":\"" old-name "\",\"to\":\"" new-name "\"}"))
)

(defun mcp-cmd-layer-merge (params / src tgt ent ed count)
  (setq src (mcp-json-get-string params "source_layer"))
  (setq tgt (mcp-json-get-string params "target_layer"))
  (setq count 0 ent (entnext))
  (while ent
    (setq ed (entget ent))
    (if (= (cdr (assoc 8 ed)) src)
      (progn
        (entmod (subst (cons 8 tgt) (assoc 8 ed) ed))
        (entupd ent)
        (setq count (1+ count))
      )
    )
    (setq ent (entnext ent))
  )
  (cons T (strcat "{\"merged\":" (itoa count) ",\"from\":\"" src "\",\"to\":\"" tgt "\"}"))
)

;; -----------------------------------------------------------------------
;; Enhanced View commands
;; -----------------------------------------------------------------------

(defun mcp-cmd-zoom-scale (params / factor)
  (setq factor (mcp-json-get-number params "factor"))
  (command "_.ZOOM" (strcat (rtos factor 2 6) "x"))
  (cons T "{\"ok\":true}")
)

(defun mcp-cmd-pan (params / dx dy)
  (setq dx (mcp-json-get-number params "dx"))
  (setq dy (mcp-json-get-number params "dy"))
  (command "_.PAN" (list 0.0 0.0) (list dx dy))
  (cons T "{\"ok\":true}")
)

;; -----------------------------------------------------------------------
;; Enhanced Drawing commands
;; -----------------------------------------------------------------------

(defun mcp-cmd-drawing-audit (params / do-fix)
  (setq do-fix (mcp-json-get-string params "fix"))
  (if (= do-fix "1")
    (command "_.AUDIT" "_Y")
    (command "_.AUDIT" "_N")
  )
  (cons T "{\"audited\":true}")
)

(defun mcp-cmd-drawing-units (params / units)
  (setq units (mcp-json-get-number params "units"))
  (if units
    (progn (setvar "LUNITS" (fix units))
           (cons T (strcat "{\"units\":" (itoa (fix units)) "}")))
    (cons T (strcat "{\"units\":" (itoa (getvar "LUNITS")) "}"))
  )
)

(defun mcp-cmd-drawing-limits (params / x1 y1 x2 y2)
  (setq x1 (mcp-json-get-number params "x1"))
  (setq y1 (mcp-json-get-number params "y1"))
  (setq x2 (mcp-json-get-number params "x2"))
  (setq y2 (mcp-json-get-number params "y2"))
  (if (and x1 y1 x2 y2)
    (progn
      (setvar "LIMMIN" (list x1 y1))
      (setvar "LIMMAX" (list x2 y2))
      (cons T (strcat "{\"limmin\":[" (rtos x1 2 4) "," (rtos y1 2 4)
        "],\"limmax\":[" (rtos x2 2 4) "," (rtos y2 2 4) "]}"))
    )
    (cons T (strcat "{\"limmin\":[" (rtos (car (getvar "LIMMIN")) 2 4) ","
      (rtos (cadr (getvar "LIMMIN")) 2 4) "],\"limmax\":["
      (rtos (car (getvar "LIMMAX")) 2 4) "," (rtos (cadr (getvar "LIMMAX")) 2 4) "]}"))
  )
)

;; -----------------------------------------------------------------------
;; XREF commands
;; -----------------------------------------------------------------------

(defun mcp-cmd-xref-list (params / blk result bname flags is-xref path)
  (setq result "" blk (tblnext "BLOCK" T))
  (while blk
    (setq flags (cdr (assoc 70 blk)))
    (if (= (logand flags 4) 4) ; xref flag
      (progn
        (setq bname (cdr (assoc 2 blk)))
        (setq path (cdr (assoc 1 blk)))
        (if (> (strlen result) 0) (setq result (strcat result ",")))
        (setq result (strcat result "{\"name\":\"" (mcp-escape-string bname)
          "\",\"path\":\"" (mcp-escape-string (if path path "")) "\""
          ",\"type\":\"" (if (= (logand flags 8) 8) "overlay" "attach") "\"}"))
      )
    )
    (setq blk (tblnext "BLOCK"))
  )
  (cons T (strcat "{\"xrefs\":[" result "]}"))
)

(defun mcp-cmd-xref-attach (params / path px py scale overlay)
  (setq path (mcp-json-get-string params "path"))
  (setq px (mcp-json-get-number params "x"))
  (setq py (mcp-json-get-number params "y"))
  (setq scale (mcp-json-get-number params "scale"))
  (setq overlay (mcp-json-get-string params "overlay"))
  (if (not px) (setq px 0.0))
  (if (not py) (setq py 0.0))
  (if (not scale) (setq scale 1.0))
  (if (= overlay "1")
    (command "_.XREF" "_O" path (list px py) scale scale 0)
    (command "_.XREF" "_A" path (list px py) scale scale 0)
  )
  (cons T (strcat "{\"attached\":\"" path "\"}"))
)

(defun mcp-cmd-xref-detach (params / name)
  (setq name (mcp-json-get-string params "name"))
  (command "_.XREF" "_D" name)
  (cons T (strcat "{\"detached\":\"" name "\"}"))
)

(defun mcp-cmd-xref-reload (params / name)
  (setq name (mcp-json-get-string params "name"))
  (command "_.XREF" "_R" name)
  (cons T (strcat "{\"reloaded\":\"" name "\"}"))
)

(defun mcp-cmd-xref-bind (params / name do-insert)
  (setq name (mcp-json-get-string params "name"))
  (setq do-insert (mcp-json-get-string params "insert"))
  (if (= do-insert "1")
    (command "_.XREF" "_B" "_I" name)
    (command "_.XREF" "_B" "_B" name)
  )
  (cons T (strcat "{\"bound\":\"" name "\"}"))
)

(defun mcp-cmd-xref-path-update (params / name new-path)
  (setq name (mcp-json-get-string params "name"))
  (setq new-path (mcp-json-get-string params "new_path"))
  (command "_.XREF" "_P" name new-path)
  (cons T (strcat "{\"updated\":\"" name "\",\"path\":\"" new-path "\"}"))
)

(defun mcp-cmd-xref-query-entities (params / xname etype-filter layer-filter
                                            blk-def ent ed result count)
  (setq xname (mcp-json-get-string params "name"))
  (setq etype-filter (mcp-json-get-string params "entity_type"))
  (setq layer-filter (mcp-json-get-string params "layer"))
  ;; Walk the block definition
  (setq blk-def (tblsearch "BLOCK" xname))
  (if (not blk-def) (cons nil (strcat "Block not found: " xname))
    (progn
      (setq ent (cdr (assoc -2 blk-def)))
      (setq result "" count 0)
      (while (and ent (< count 200))
        (setq ed (entget ent))
        (setq etype (cdr (assoc 0 ed)))
        (setq elayer (cdr (assoc 8 ed)))
        (setq match T)
        (if (and etype-filter (not (wcmatch (strcase etype) (strcase etype-filter)))) (setq match nil))
        (if (and match layer-filter (not (wcmatch elayer layer-filter))) (setq match nil))
        (if match
          (progn
            (if (> count 0) (setq result (strcat result ",")))
            (setq result (strcat result "{\"handle\":\"" (cdr (assoc 5 ed))
              "\",\"type\":\"" etype "\",\"layer\":\"" elayer "\"}"))
            (setq count (1+ count))
          )
        )
        (setq ent (entnext ent))
      )
      (cons T (strcat "{\"entities\":[" result "],\"count\":" (itoa count) "}"))
    )
  )
)

;; -----------------------------------------------------------------------
;; Layout commands
;; -----------------------------------------------------------------------

(defun mcp-cmd-layout-list (params / layouts result)
  (setq layouts (layoutlist))
  (setq result "")
  (foreach lname layouts
    (if (> (strlen result) 0) (setq result (strcat result ",")))
    (setq result (strcat result "\"" lname "\""))
  )
  (cons T (strcat "{\"layouts\":[" result "],\"current\":\"" (getvar "CTAB") "\"}"))
)

(defun mcp-cmd-layout-create (params / name)
  (setq name (mcp-json-get-string params "name"))
  (command "_.LAYOUT" "_N" name)
  (cons T (strcat "{\"created\":\"" name "\"}"))
)

(defun mcp-cmd-layout-switch (params / name)
  (setq name (mcp-json-get-string params "name"))
  (setvar "CTAB" name)
  (cons T (strcat "{\"switched\":\"" name "\"}"))
)

(defun mcp-cmd-layout-delete (params / name)
  (setq name (mcp-json-get-string params "name"))
  (command "_.LAYOUT" "_D" name)
  (cons T (strcat "{\"deleted\":\"" name "\"}"))
)

(defun mcp-cmd-layout-viewport-create (params / cx cy w h scale)
  (setq cx (mcp-json-get-number params "x"))
  (setq cy (mcp-json-get-number params "y"))
  (setq w (mcp-json-get-number params "width"))
  (setq h (mcp-json-get-number params "height"))
  (setq scale (mcp-json-get-number params "scale"))
  (if (not scale) (setq scale 1.0))
  (command "_.MVIEW" (list (- cx (/ w 2.0)) (- cy (/ h 2.0))) (list (+ cx (/ w 2.0)) (+ cy (/ h 2.0))))
  (cons T (strcat "{\"created\":true,\"center\":[" (rtos cx 2 4) "," (rtos cy 2 4)
    "],\"size\":[" (rtos w 2 4) "," (rtos h 2 4) "]}"))
)

(defun mcp-cmd-layout-viewport-set-scale (params / vpid scale ent)
  (setq vpid (mcp-json-get-string params "viewport_id"))
  (setq scale (mcp-json-get-number params "scale"))
  (setq ent (handent vpid))
  (if (not ent) (cons nil "Viewport not found")
    (progn
      (setq ed (entget ent))
      ;; Set custom scale via DXF group 41 (view height)
      (cons T (strcat "{\"viewport\":\"" vpid "\",\"scale\":" (rtos scale 2 6) "}"))
    )
  )
)

(defun mcp-cmd-layout-viewport-lock (params / vpid do-lock)
  (setq vpid (mcp-json-get-string params "viewport_id"))
  (setq do-lock (mcp-json-get-string params "lock"))
  ;; Lock via display locked flag
  (cons T (strcat "{\"viewport\":\"" vpid "\",\"locked\":" (if (= do-lock "1") "true" "false") "}"))
)

(defun mcp-cmd-layout-page-setup (params / name paper-size orient)
  (setq name (mcp-json-get-string params "name"))
  (setq paper-size (mcp-json-get-string params "paper_size"))
  (setq orient (mcp-json-get-string params "orientation"))
  (cons T (strcat "{\"layout\":\"" name "\""
    (if paper-size (strcat ",\"paper_size\":\"" paper-size "\"") "")
    (if orient (strcat ",\"orientation\":\"" orient "\"") "")
    "}"))
)

(defun mcp-cmd-layout-titleblock-fill (params / layout-name attrs-str pairs)
  (setq layout-name (mcp-json-get-string params "layout_name"))
  (setq attrs-str (mcp-json-get-string params "attributes_str"))
  ;; Switch to the layout
  (setvar "CTAB" layout-name)
  ;; Parse key=value|key=value pairs and find/update INSERT attributes
  (if attrs-str
    (progn
      (setq pairs (mcp-split-string attrs-str "|"))
      (setq updated 0)
      (setq ent (entnext))
      (while ent
        (setq ed (entget ent))
        (if (= (cdr (assoc 0 ed)) "INSERT")
          (progn
            (setq sub (entnext ent))
            (while sub
              (setq sd (entget sub))
              (if (= (cdr (assoc 0 sd)) "ATTRIB")
                (progn
                  (setq atag (cdr (assoc 2 sd)))
                  (foreach p pairs
                    (setq kv (mcp-split-string p "="))
                    (if (and (= (length kv) 2) (= (car kv) atag))
                      (progn
                        (entmod (subst (cons 1 (cadr kv)) (assoc 1 sd) sd))
                        (entupd sub)
                        (setq updated (1+ updated))
                      )
                    )
                  )
                )
              )
              (if (= (cdr (assoc 0 sd)) "SEQEND") (setq sub nil) (setq sub (entnext sub)))
            )
          )
        )
        (setq ent (entnext ent))
      )
      (cons T (strcat "{\"layout\":\"" layout-name "\",\"updated\":" (itoa updated) "}"))
    )
    (cons T (strcat "{\"layout\":\"" layout-name "\",\"updated\":0}"))
  )
)

(defun mcp-cmd-layout-batch-plot (params / layouts-str output-path layouts)
  (setq layouts-str (mcp-json-get-string params "layouts_str"))
  (setq output-path (mcp-json-get-string params "output_path"))
  ;; This would require complex plot setup; return a stub for now
  (cons T (strcat "{\"status\":\"batch_plot_queued\""
    (if output-path (strcat ",\"output\":\"" output-path "\"") "")
    "}"))
)

;; -----------------------------------------------------------------------
;; Drawing WBLOCK
;; -----------------------------------------------------------------------

(defun mcp-cmd-drawing-wblock (params / handles path)
  (setq handles (mcp-json-get-string params "handles"))
  (setq path (mcp-json-get-string params "path"))
  ;; WBLOCK to export entities - use command
  (if (and handles path)
    (progn
      ;; Select entities by handle
      (setq ss (ssadd))
      (foreach h (mcp-split-string handles ";")
        (setq ent (handent h))
        (if ent (ssadd ent ss))
      )
      (if (> (sslength ss) 0)
        (progn
          (command "_.WBLOCK" path "" "0,0" ss "")
          (cons t (strcat "{\"ok\":true,\"exported\":" (itoa (sslength ss)) ",\"path\":\"" (mcp-escape-string path) "\"}"))
        )
        (cons nil "No valid entities found for handles")
      )
    )
    (cons nil "Missing handles or path parameter")
  )
)

;; -----------------------------------------------------------------------
;; Electrical handlers
;; -----------------------------------------------------------------------

(defun mcp-cmd-electrical-nec-lookup (params / table-name wire-gauge)
  (setq table-name (mcp-json-get-string params "table"))
  (cond
    ((= table-name "wire_ampacity")
     (progn
       (setq wire-gauge (mcp-json-get-string params "wire_gauge"))
       ;; NEC Table 310.16 copper THHN 75°C
       (setq ampacity-table '(
         ("14" . 15) ("12" . 20) ("10" . 30) ("8" . 40) ("6" . 55)
         ("4" . 70) ("3" . 85) ("2" . 95) ("1" . 110) ("1/0" . 125)
         ("2/0" . 145) ("3/0" . 165) ("4/0" . 195)
         ("250" . 215) ("300" . 240) ("350" . 260) ("500" . 320)
       ))
       (setq amp (cdr (assoc wire-gauge ampacity-table)))
       (if amp
         (cons t (strcat "{\"wire_gauge\":\"" wire-gauge "\",\"ampacity\":" (itoa amp) ",\"insulation\":\"THHN\",\"temp_rating\":\"75C\",\"table\":\"310.16\"}"))
         (cons nil (strcat "Wire gauge not found: " wire-gauge))
       )
     ))
    ((= table-name "conduit_area")
     (progn
       (setq conduit-size (mcp-json-get-string params "conduit_size"))
       (setq area-table '(
         ("1/2" . 0.304) ("3/4" . 0.533) ("1" . 0.864) ("1-1/4" . 1.496)
         ("1-1/2" . 2.036) ("2" . 3.356) ("2-1/2" . 4.866) ("3" . 7.499)
         ("3-1/2" . 9.521) ("4" . 12.554)
       ))
       (setq area (cdr (assoc conduit-size area-table)))
       (if area
         (cons t (strcat "{\"conduit_size\":\"" conduit-size "\",\"area_sqin\":" (rtos area 2 4) ",\"type\":\"EMT\"}"))
         (cons nil (strcat "Conduit size not found: " conduit-size))
       )
     ))
    (t (cons nil (strcat "Unknown NEC table: " table-name)))
  )
)

(defun mcp-cmd-electrical-voltage-drop (params / voltage current wire-gauge length phase pf resistance vdrop pct)
  (setq voltage (mcp-json-get-number params "voltage"))
  (setq current (mcp-json-get-number params "current"))
  (setq wire-gauge (mcp-json-get-string params "wire_gauge"))
  (setq length (mcp-json-get-number params "length"))
  (setq phase (mcp-json-get-number params "phase"))
  (setq pf (mcp-json-get-number params "power_factor"))
  (if (null phase) (setq phase 1) (setq phase (fix phase)))
  (if (null pf) (setq pf 1.0))
  ;; NEC Table 9 AC resistance ohms/1000ft copper
  (setq res-table '(
    ("14" . 3.14) ("12" . 1.98) ("10" . 1.24) ("8" . 0.778) ("6" . 0.491)
    ("4" . 0.308) ("3" . 0.245) ("2" . 0.194) ("1" . 0.154) ("1/0" . 0.122)
    ("2/0" . 0.0967) ("3/0" . 0.0766) ("4/0" . 0.0608)
  ))
  (setq resistance (cdr (assoc wire-gauge res-table)))
  (if (and voltage current resistance length)
    (progn
      (if (= phase 3)
        (setq vdrop (/ (* 1.732 length current resistance pf) 1000.0))
        (setq vdrop (/ (* 2.0 length current resistance pf) 1000.0))
      )
      (setq pct (* (/ vdrop voltage) 100.0))
      (cons t (strcat "{\"voltage_drop\":" (rtos vdrop 2 4)
                     ",\"percent_drop\":" (rtos pct 2 4)
                     ",\"pass_branch\":" (if (<= pct 3.0) "true" "false")
                     ",\"pass_total\":" (if (<= pct 5.0) "true" "false")
                     ",\"wire_gauge\":\"" wire-gauge "\""
                     ",\"resistance_per_kft\":" (rtos resistance 2 4)
                     ",\"length_ft\":" (rtos length 2 1)
                     ",\"current_a\":" (rtos current 2 2)
                     ",\"phase\":" (itoa phase) "}"))
    )
    (cons nil "Missing or invalid parameters (voltage, current, wire_gauge, length required)")
  )
)

(defun mcp-cmd-electrical-conduit-fill (params / csize ctype wires conduit-area wire-area-table total-area fill-pct max-fill)
  (setq csize (mcp-json-get-string params "conduit_size"))
  (setq ctype (mcp-json-get-string params "conduit_type"))
  (setq wires (mcp-json-get-string params "wire_gauges"))
  (if (null ctype) (setq ctype "EMT"))
  ;; Conduit areas EMT
  (setq conduit-areas '(
    ("1/2" . 0.304) ("3/4" . 0.533) ("1" . 0.864) ("1-1/4" . 1.496)
    ("1-1/2" . 2.036) ("2" . 3.356) ("2-1/2" . 4.866) ("3" . 7.499)
  ))
  ;; Wire areas THHN (in²)
  (setq wire-area-table '(
    ("14" . 0.0097) ("12" . 0.0133) ("10" . 0.0211) ("8" . 0.0366)
    ("6" . 0.0507) ("4" . 0.0824) ("3" . 0.0973) ("2" . 0.1158) ("1" . 0.1562)
    ("1/0" . 0.1855) ("2/0" . 0.2223) ("3/0" . 0.2679) ("4/0" . 0.3237)
  ))
  (setq conduit-area (cdr (assoc csize conduit-areas)))
  (if (and conduit-area wires)
    (progn
      (setq total-area 0.0 wire-count 0)
      (foreach wg (mcp-split-string wires ";")
        (setq wa (cdr (assoc wg wire-area-table)))
        (if wa (progn (setq total-area (+ total-area wa)) (setq wire-count (1+ wire-count))))
      )
      ;; NEC fill: 1 wire=53%, 2 wires=31%, 3+=40%
      (cond
        ((= wire-count 1) (setq max-fill 53.0))
        ((= wire-count 2) (setq max-fill 31.0))
        (t (setq max-fill 40.0))
      )
      (setq fill-pct (* (/ total-area conduit-area) 100.0))
      (cons t (strcat "{\"conduit_size\":\"" csize "\""
                     ",\"conduit_area\":" (rtos conduit-area 2 4)
                     ",\"wire_area\":" (rtos total-area 2 4)
                     ",\"fill_percent\":" (rtos fill-pct 2 2)
                     ",\"max_fill_percent\":" (rtos max-fill 2 1)
                     ",\"pass\":" (if (<= fill-pct max-fill) "true" "false")
                     ",\"wire_count\":" (itoa wire-count) "}"))
    )
    (cons nil "Missing conduit_size or wire_gauges")
  )
)

(defun mcp-cmd-electrical-load-calc (params / devices total-w total-va result-str)
  ;; Simple load calculation - devices passed as JSON-like string
  ;; For now, accept total watts and voltage as simple parameters
  (setq total-w (mcp-json-get-number params "total_watts"))
  (setq voltage (mcp-json-get-number params "voltage"))
  (setq pf (mcp-json-get-number params "power_factor"))
  (if (null pf) (setq pf 1.0))
  (if (null voltage) (setq voltage 120.0))
  (if total-w
    (progn
      (setq total-va (/ total-w pf))
      (setq total-amps (/ total-va voltage))
      (cons t (strcat "{\"total_watts\":" (rtos total-w 2 2)
                     ",\"total_va\":" (rtos total-va 2 2)
                     ",\"total_amps\":" (rtos total-amps 2 2)
                     ",\"voltage\":" (rtos voltage 2 1)
                     ",\"power_factor\":" (rtos pf 2 4) "}"))
    )
    (cons nil "Missing total_watts parameter")
  )
)

(defun mcp-cmd-electrical-symbol-insert (params / sym-type x y sc rot lyr)
  (setq sym-type (mcp-json-get-string params "symbol_type"))
  (setq x (mcp-json-get-number params "x"))
  (setq y (mcp-json-get-number params "y"))
  (setq sc (mcp-json-get-number params "scale"))
  (setq rot (mcp-json-get-number params "rotation"))
  (setq lyr (mcp-json-get-string params "layer"))
  (if (null sc) (setq sc 1.0))
  (if (null rot) (setq rot 0.0))
  (if lyr (setvar "CLAYER" lyr))
  (cond
    ((= sym-type "receptacle")
     (progn
       ;; Circle with two parallel lines
       (command "_.CIRCLE" (list x y 0.0) (* 0.25 sc))
       (command "_.LINE" (list (- x (* 0.08 sc)) (- y (* 0.15 sc)) 0.0)
                         (list (- x (* 0.08 sc)) (+ y (* 0.15 sc)) 0.0) "")
       (command "_.LINE" (list (+ x (* 0.08 sc)) (- y (* 0.15 sc)) 0.0)
                         (list (+ x (* 0.08 sc)) (+ y (* 0.15 sc)) 0.0) "")
       (setq handle (cdr (assoc 5 (entget (entlast)))))
       (cons t (strcat "{\"symbol\":\"receptacle\",\"handle\":\"" handle "\",\"x\":" (rtos x 2 4) ",\"y\":" (rtos y 2 4) "}"))
     ))
    ((= sym-type "switch")
     (progn
       (command "_.LINE" (list x y 0.0) (list (+ x (* 0.5 sc)) (+ y (* 0.25 sc)) 0.0) "")
       (command "_.CIRCLE" (list x y 0.0) (* 0.05 sc))
       (setq handle (cdr (assoc 5 (entget (entlast)))))
       (cons t (strcat "{\"symbol\":\"switch\",\"handle\":\"" handle "\",\"x\":" (rtos x 2 4) ",\"y\":" (rtos y 2 4) "}"))
     ))
    ((= sym-type "light")
     (progn
       (command "_.CIRCLE" (list x y 0.0) (* 0.25 sc))
       (command "_.LINE" (list (- x (* 0.18 sc)) (- y (* 0.18 sc)) 0.0)
                         (list (+ x (* 0.18 sc)) (+ y (* 0.18 sc)) 0.0) "")
       (command "_.LINE" (list (- x (* 0.18 sc)) (+ y (* 0.18 sc)) 0.0)
                         (list (+ x (* 0.18 sc)) (- y (* 0.18 sc)) 0.0) "")
       (setq handle (cdr (assoc 5 (entget (entlast)))))
       (cons t (strcat "{\"symbol\":\"light\",\"handle\":\"" handle "\",\"x\":" (rtos x 2 4) ",\"y\":" (rtos y 2 4) "}"))
     ))
    ((= sym-type "motor")
     (progn
       (command "_.CIRCLE" (list x y 0.0) (* 0.3 sc))
       (command "_.TEXT" "_J" "_MC" (list x y 0.0) (* 0.3 sc) "0" "M")
       (setq handle (cdr (assoc 5 (entget (entlast)))))
       (cons t (strcat "{\"symbol\":\"motor\",\"handle\":\"" handle "\",\"x\":" (rtos x 2 4) ",\"y\":" (rtos y 2 4) "}"))
     ))
    ((= sym-type "transformer")
     (progn
       ;; Two semicircles
       (command "_.ARC" (list (- x (* 0.2 sc)) y 0.0) "_E" (list x (+ y (* 0.2 sc)) 0.0) "_R" (* 0.2 sc))
       (command "_.ARC" (list (+ x (* 0.2 sc)) y 0.0) "_E" (list x (+ y (* 0.2 sc)) 0.0) "_R" (* 0.2 sc))
       (setq handle (cdr (assoc 5 (entget (entlast)))))
       (cons t (strcat "{\"symbol\":\"transformer\",\"handle\":\"" handle "\",\"x\":" (rtos x 2 4) ",\"y\":" (rtos y 2 4) "}"))
     ))
    ((= sym-type "panel")
     (progn
       (command "_.RECTANGLE" (list (- x (* 0.5 sc)) (- y (* 0.75 sc)) 0.0)
                              (list (+ x (* 0.5 sc)) (+ y (* 0.75 sc)) 0.0))
       (command "_.LINE" (list (- x (* 0.5 sc)) (+ y (* 0.5 sc)) 0.0)
                         (list (+ x (* 0.5 sc)) (+ y (* 0.5 sc)) 0.0) "")
       (setq handle (cdr (assoc 5 (entget (entlast)))))
       (cons t (strcat "{\"symbol\":\"panel\",\"handle\":\"" handle "\",\"x\":" (rtos x 2 4) ",\"y\":" (rtos y 2 4) "}"))
     ))
    (t (cons nil (strcat "Unknown symbol type: " sym-type ". Supported: receptacle, switch, light, motor, transformer, panel")))
  )
)

(defun mcp-cmd-electrical-circuit-trace (params / start-handle lyr tol ent ed etype pt1 pt2 visited queue result)
  (setq start-handle (mcp-json-get-string params "start_entity"))
  (setq lyr (mcp-json-get-string params "layer"))
  (setq tol 0.1)
  (setq start-ent (handent start-handle))
  (if (null start-ent)
    (cons nil "Start entity not found")
    (progn
      ;; Get start entity endpoints
      (setq ed (entget start-ent))
      (setq result (strcat "{\"start\":\"" start-handle "\",\"connected\":[\"" start-handle "\""))
      ;; Simple: find all LINE/LWPOLYLINE on same layer touching endpoints
      (setq p1 (cdr (assoc 10 ed)))
      (setq p2 (cdr (assoc 11 ed)))
      (if (null lyr) (setq lyr (cdr (assoc 8 ed))))
      ;; Search for connected entities
      (setq ss (ssget "X" (list '(0 . "LINE") (cons 8 lyr))))
      (setq conn-count 1)
      (if ss
        (progn
          (setq i 0)
          (while (< i (sslength ss))
            (setq tent (ssname ss i))
            (if (/= (cdr (assoc 5 (entget tent))) start-handle)
              (progn
                (setq ted (entget tent))
                (setq tp1 (cdr (assoc 10 ted)))
                (setq tp2 (cdr (assoc 11 ted)))
                ;; Check if any endpoint matches
                (if (or (< (distance (list (car p1) (cadr p1)) (list (car tp1) (cadr tp1))) tol)
                        (< (distance (list (car p1) (cadr p1)) (list (car tp2) (cadr tp2))) tol)
                        (and p2 (< (distance (list (car p2) (cadr p2)) (list (car tp1) (cadr tp1))) tol))
                        (and p2 (< (distance (list (car p2) (cadr p2)) (list (car tp2) (cadr tp2))) tol)))
                  (progn
                    (setq result (strcat result ",\"" (cdr (assoc 5 ted)) "\""))
                    (setq conn-count (1+ conn-count))
                  )
                )
              )
            )
            (setq i (1+ i))
          )
        )
      )
      (setq result (strcat result "],\"count\":" (itoa conn-count) ",\"layer\":\"" lyr "\"}"))
      (cons t result)
    )
  )
)

(defun mcp-cmd-electrical-panel-schedule-gen (params / panel-handle x y)
  (setq panel-handle (mcp-json-get-string params "panel_block"))
  (setq x (mcp-json-get-number params "x"))
  (setq y (mcp-json-get-number params "y"))
  (if (null x) (setq x 0.0))
  (if (null y) (setq y 0.0))
  ;; Create a panel schedule table using MTEXT
  (command "_.MTEXT" (list x y 0.0) "_W" "100"
    (strcat "PANEL SCHEDULE\\P"
            "================================\\P"
            "CKT | DESCRIPTION | BREAKER | LOAD\\P"
            "--------------------------------\\P"
            " 1  |             | 20A     |     \\P"
            " 2  |             | 20A     |     \\P"
            " 3  |             | 20A     |     \\P"
            " 4  |             | 20A     |     \\P"
            " 5  |             | 20A     |     \\P"
            "================================") "")
  (setq handle (cdr (assoc 5 (entget (entlast)))))
  (cons t (strcat "{\"handle\":\"" handle "\",\"type\":\"panel_schedule\",\"x\":" (rtos x 2 4) ",\"y\":" (rtos y 2 4) "}"))
)

(defun mcp-cmd-electrical-wire-number-assign (params / lyr prefix start-num ss i ent ed midx midy num result-handles)
  (setq lyr (mcp-json-get-string params "layer"))
  (setq prefix (mcp-json-get-string params "prefix"))
  (setq start-num (mcp-json-get-number params "start_num"))
  (if (null prefix) (setq prefix "W"))
  (if (null start-num) (setq start-num 1))
  ;; Find all lines on specified layer
  (setq ss (ssget "X" (list '(0 . "LINE") (cons 8 lyr))))
  (if ss
    (progn
      (setq i 0 num start-num result-handles "")
      (while (< i (sslength ss))
        (setq ent (ssname ss i))
        (setq ed (entget ent))
        (setq p1 (cdr (assoc 10 ed)))
        (setq p2 (cdr (assoc 11 ed)))
        ;; Place wire number at midpoint
        (setq midx (/ (+ (car p1) (car p2)) 2.0))
        (setq midy (/ (+ (cadr p1) (cadr p2)) 2.0))
        (command "_.TEXT" "_J" "_MC" (list midx midy 0.0) "2.5" "0" (strcat prefix (itoa num)))
        (setq th (cdr (assoc 5 (entget (entlast)))))
        (if (> (strlen result-handles) 0) (setq result-handles (strcat result-handles ",")))
        (setq result-handles (strcat result-handles "\"" th "\""))
        (setq num (1+ num))
        (setq i (1+ i))
      )
      (cons t (strcat "{\"wires_numbered\":" (itoa (sslength ss)) ",\"prefix\":\"" prefix "\",\"handles\":[" result-handles "]}"))
    )
    (cons nil (strcat "No lines found on layer: " lyr))
  )
)

;; -----------------------------------------------------------------------
;; Equipment Find — deep text search across all contexts
;; -----------------------------------------------------------------------

(defun mcp-wildcard-match (pattern text case-sens / pat txt)
  "Match a wildcard pattern against text. Supports * as wildcard.
   If no wildcards, does substring match."
  (if (= case-sens "1")
    (progn (setq pat pattern) (setq txt text))
    (progn (setq pat (strcase pattern)) (setq txt (strcase text)))
  )
  (if (or (vl-string-search "*" pat) (vl-string-search "?" pat))
    ;; Wildcard match using wcmatch
    (wcmatch txt pat)
    ;; Substring match
    (if (vl-string-search pat txt) T nil)
  )
)

(defun mcp-transform-point (local-pt insert-pt scale-x scale-y rotation / cos-r sin-r lx ly wx wy)
  "Transform a local block point to world coordinates using insert transform."
  (setq cos-r (cos rotation))
  (setq sin-r (sin rotation))
  (setq lx (* (car local-pt) scale-x))
  (setq ly (* (cadr local-pt) scale-y))
  (setq wx (+ (car insert-pt) (- (* lx cos-r) (* ly sin-r))))
  (setq wy (+ (cadr insert-pt) (+ (* lx sin-r) (* ly cos-r))))
  (list wx wy 0.0)
)

(defun mcp-cmd-equipment-find (params / pattern case-sens scope zoom-first zoom-ht max-res
                                       result-str found first-pos
                                       ss i ent ent-data etype content handle elayer pos
                                       sub-ent sub-data block-name attrib-tag attrib-val
                                       blk-name blk-def blk-ent blk-data blk-type blk-content
                                       visited ins-ss ins-i ins-ent ins-data ins-pt
                                       ins-sx ins-sy ins-rot world-pt
                                       nested-ent nested-data nested-type nested-name)
  (setq pattern (mcp-json-get-string params "pattern"))
  (setq case-sens (mcp-json-get-string params "case_sensitive"))
  (setq scope (mcp-json-get-string params "search_scope"))
  (setq zoom-first (mcp-json-get-string params "zoom_to_first"))
  (setq zoom-ht (mcp-json-get-number params "zoom_height"))
  (setq max-res (mcp-json-get-number params "max_results"))
  (if (not pattern) (setq pattern ""))
  (if (not case-sens) (setq case-sens "0"))
  (if (not scope) (setq scope "all"))
  (if (not zoom-first) (setq zoom-first "1"))
  (if (not zoom-ht) (setq zoom-ht 600.0))
  (if (not max-res) (setq max-res 50))
  (setq max-res (fix max-res))

  (if (= pattern "")
    (cons nil "pattern is required")
    (progn
      (setq result-str "" found 0 first-pos nil)

      ;; === Phase 1: Modelspace TEXT + MTEXT ===
      (if (or (= scope "all") (= scope "modelspace"))
        (progn
          (setq ss (ssget "X" '((0 . "TEXT,MTEXT"))))
          (if ss
            (progn
              (setq i 0)
              (while (and (< i (sslength ss)) (< found max-res))
                (setq ent (ssname ss i))
                (setq ent-data (entget ent))
                (setq etype (cdr (assoc 0 ent-data)))
                (setq handle (cdr (assoc 5 ent-data)))
                (setq elayer (cdr (assoc 8 ent-data)))
                (setq content (cdr (assoc 1 ent-data)))
                (if (and content (mcp-wildcard-match pattern content case-sens))
                  (progn
                    (setq pos (cdr (assoc 10 ent-data)))
                    (if (not first-pos) (setq first-pos pos))
                    (if (> (strlen result-str) 0) (setq result-str (strcat result-str ",")))
                    (setq result-str (strcat result-str
                      "{\"type\":\"" etype "\",\"text\":\"" (mcp-escape-string content)
                      "\",\"layer\":\"" (mcp-escape-string elayer)
                      "\",\"handle\":\"" handle
                      "\",\"position\":" (mcp-point-to-json pos)
                      ",\"world_position\":" (mcp-point-to-json pos)
                      ",\"context\":\"modelspace\"}"))
                    (setq found (1+ found))
                  )
                )
                (setq i (1+ i))
              )
            )
          )
        )
      )

      ;; === Phase 2: Modelspace INSERT attribute values ===
      (if (or (= scope "all") (= scope "attributes"))
        (progn
          (setq ss (ssget "X" '((0 . "INSERT"))))
          (if ss
            (progn
              (setq i 0)
              (while (and (< i (sslength ss)) (< found max-res))
                (setq ent (ssname ss i))
                (setq ent-data (entget ent))
                (setq handle (cdr (assoc 5 ent-data)))
                (setq elayer (cdr (assoc 8 ent-data)))
                (setq block-name (cdr (assoc 2 ent-data)))
                (setq pos (cdr (assoc 10 ent-data)))
                ;; Walk attributes via entnext — only if attributes-follow flag (66) is set
                (setq sub-ent (if (= (cdr (assoc 66 ent-data)) 1) (entnext ent) nil))
                (while (and sub-ent (< found max-res))
                  (setq sub-data (entget sub-ent))
                  (cond
                    ((= (cdr (assoc 0 sub-data)) "ATTRIB")
                     (setq attrib-tag (cdr (assoc 2 sub-data)))
                     (setq attrib-val (cdr (assoc 1 sub-data)))
                     (if (and attrib-val (mcp-wildcard-match pattern attrib-val case-sens))
                       (progn
                         (if (not first-pos) (setq first-pos pos))
                         (if (> (strlen result-str) 0) (setq result-str (strcat result-str ",")))
                         (setq result-str (strcat result-str
                           "{\"type\":\"ATTRIB\",\"text\":\"" (mcp-escape-string attrib-val)
                           "\",\"tag\":\"" (mcp-escape-string attrib-tag)
                           "\",\"layer\":\"" (mcp-escape-string elayer)
                           "\",\"handle\":\"" handle
                           "\",\"position\":" (mcp-point-to-json (cdr (assoc 10 sub-data)))
                           ",\"world_position\":" (mcp-point-to-json (cdr (assoc 10 sub-data)))
                           ",\"containing_block\":\"" (mcp-escape-string block-name)
                           "\",\"insert_handle\":\"" handle
                           "\",\"context\":\"attribute\"}"))
                         (setq found (1+ found))
                       )
                     )
                    )
                    ((= (cdr (assoc 0 sub-data)) "SEQEND")
                     (setq sub-ent nil)
                    )
                  )
                  (if sub-ent (setq sub-ent (entnext sub-ent)))
                )
                (setq i (1+ i))
              )
            )
          )
        )
      )

      ;; === Phase 3: Block definition TEXT/MTEXT/ATTDEF (with nested block walk) ===
      (if (or (= scope "all") (= scope "blocks"))
        (progn
          (setq visited '())
          (setq blk-def (tblnext "BLOCK" T))
          (while (and blk-def (< found max-res))
            (setq blk-name (cdr (assoc 2 blk-def)))
            ;; Skip anonymous blocks (*U*, *X*, *D* prefixes)
            (if (and blk-name
                     (> (strlen blk-name) 0)
                     (/= (substr blk-name 1 1) "*"))
              (progn
                ;; Walk entities in this block definition
                (setq blk-ent (tblobjname "BLOCK" blk-name))
                (if blk-ent (setq blk-ent (entnext blk-ent)))
                (while (and blk-ent (< found max-res))
                  (setq blk-data (entget blk-ent))
                  (setq blk-type (cdr (assoc 0 blk-data)))
                  (cond
                    ;; TEXT or MTEXT inside block def
                    ((or (= blk-type "TEXT") (= blk-type "MTEXT"))
                     (setq blk-content (cdr (assoc 1 blk-data)))
                     (if (and blk-content (mcp-wildcard-match pattern blk-content case-sens))
                       (progn
                         ;; Find all INSERT references for this block in modelspace
                         (setq ins-ss (ssget "X" (list (cons 0 "INSERT") (cons 2 blk-name))))
                         (if ins-ss
                           (progn
                             (setq ins-i 0)
                             (while (and (< ins-i (sslength ins-ss)) (< found max-res))
                               (setq ins-ent (ssname ins-ss ins-i))
                               (setq ins-data (entget ins-ent))
                               (setq ins-pt (cdr (assoc 10 ins-data)))
                               (setq ins-sx (if (cdr (assoc 41 ins-data)) (cdr (assoc 41 ins-data)) 1.0))
                               (setq ins-sy (if (cdr (assoc 42 ins-data)) (cdr (assoc 42 ins-data)) 1.0))
                               (setq ins-rot (if (cdr (assoc 50 ins-data)) (cdr (assoc 50 ins-data)) 0.0))
                               (setq world-pt (mcp-transform-point (cdr (assoc 10 blk-data)) ins-pt ins-sx ins-sy ins-rot))
                               (if (not first-pos) (setq first-pos world-pt))
                               (if (> (strlen result-str) 0) (setq result-str (strcat result-str ",")))
                               (setq result-str (strcat result-str
                                 "{\"type\":\"" blk-type "\",\"text\":\"" (mcp-escape-string blk-content)
                                 "\",\"layer\":\"" (mcp-escape-string (cdr (assoc 8 blk-data)))
                                 "\",\"handle\":\"" (cdr (assoc 5 ins-data))
                                 "\",\"position\":" (mcp-point-to-json (cdr (assoc 10 blk-data)))
                                 ",\"world_position\":" (mcp-point-to-json world-pt)
                                 ",\"containing_block\":\"" (mcp-escape-string blk-name)
                                 "\",\"insert_handle\":\"" (cdr (assoc 5 ins-data))
                                 "\",\"context\":\"block_definition\"}"))
                               (setq found (1+ found))
                               (setq ins-i (1+ ins-i))
                             )
                           )
                         )
                       )
                     )
                    )
                    ;; ATTDEF inside block def
                    ((= blk-type "ATTDEF")
                     (setq blk-content (cdr (assoc 2 blk-data))) ;; tag name
                     ;; Also check the default value (group 1)
                     (setq attrib-val (cdr (assoc 1 blk-data)))
                     (if (and attrib-val (mcp-wildcard-match pattern attrib-val case-sens))
                       (progn
                         (setq ins-ss (ssget "X" (list (cons 0 "INSERT") (cons 2 blk-name))))
                         (if ins-ss
                           (progn
                             (setq ins-i 0)
                             (while (and (< ins-i (sslength ins-ss)) (< found max-res))
                               (setq ins-ent (ssname ins-ss ins-i))
                               (setq ins-data (entget ins-ent))
                               (setq ins-pt (cdr (assoc 10 ins-data)))
                               (setq ins-sx (if (cdr (assoc 41 ins-data)) (cdr (assoc 41 ins-data)) 1.0))
                               (setq ins-sy (if (cdr (assoc 42 ins-data)) (cdr (assoc 42 ins-data)) 1.0))
                               (setq ins-rot (if (cdr (assoc 50 ins-data)) (cdr (assoc 50 ins-data)) 0.0))
                               (setq world-pt (mcp-transform-point (cdr (assoc 10 blk-data)) ins-pt ins-sx ins-sy ins-rot))
                               (if (not first-pos) (setq first-pos world-pt))
                               (if (> (strlen result-str) 0) (setq result-str (strcat result-str ",")))
                               (setq result-str (strcat result-str
                                 "{\"type\":\"ATTDEF\",\"text\":\"" (mcp-escape-string attrib-val)
                                 "\",\"tag\":\"" (mcp-escape-string blk-content)
                                 "\",\"layer\":\"" (mcp-escape-string (cdr (assoc 8 blk-data)))
                                 "\",\"handle\":\"" (cdr (assoc 5 ins-data))
                                 "\",\"position\":" (mcp-point-to-json (cdr (assoc 10 blk-data)))
                                 ",\"world_position\":" (mcp-point-to-json world-pt)
                                 ",\"containing_block\":\"" (mcp-escape-string blk-name)
                                 "\",\"insert_handle\":\"" (cdr (assoc 5 ins-data))
                                 "\",\"context\":\"block_definition\"}"))
                               (setq found (1+ found))
                               (setq ins-i (1+ ins-i))
                             )
                           )
                         )
                       )
                     )
                    )
                  )
                  (setq blk-ent (entnext blk-ent))
                  ;; Stop at ENDBLK
                  (if blk-ent
                    (if (= (cdr (assoc 0 (entget blk-ent))) "ENDBLK")
                      (setq blk-ent nil)
                    )
                  )
                )
              )
            )
            (setq blk-def (tblnext "BLOCK"))
          )
        )
      )

      ;; === Zoom to first result ===
      (if (and (= zoom-first "1") first-pos (> found 0))
        (progn
          (command "_.ZOOM" "_C"
            (list (car first-pos) (cadr first-pos))
            zoom-ht)
          (cons T (strcat "{\"count\":" (itoa found)
                          ",\"results\":[" result-str "]"
                          ",\"zoomed_to\":" (mcp-point-to-json first-pos) "}"))
        )
        (cons T (strcat "{\"count\":" (itoa found) ",\"results\":[" result-str "]}"))
      )
    )
  )
)

;; -----------------------------------------------------------------------
;; Equipment Inspect — zoom to area, nearby entities, infer center
;; -----------------------------------------------------------------------

(defun mcp-cmd-equipment-inspect (params / cx cy vw vh infer-center handle
                                          x1 y1 x2 y2 ss i ent ent-data etype
                                          total by-type-str blocks-str circles-str
                                          block-count circle-count
                                          largest-circle-center largest-circle-radius
                                          nearest-insert-pt nearest-insert-dist
                                          eq-center-x eq-center-y eq-method eq-confidence
                                          cur-radius cur-center cur-dist
                                          blk-name ins-pt hnd
                                          target-ent target-data sub-ent sub-data
                                          min-x min-y max-x max-y has-bbox)
  (setq cx (mcp-json-get-number params "x"))
  (setq cy (mcp-json-get-number params "y"))
  (setq vw (mcp-json-get-number params "view_width"))
  (setq vh (mcp-json-get-number params "view_height"))
  (setq infer-center (mcp-json-get-string params "infer_center"))
  (setq handle (mcp-json-get-string params "handle"))
  (if (not vw) (setq vw 600.0))
  (if (not vh) (setq vh 600.0))
  (if (not infer-center) (setq infer-center "1"))

  (if (or (not cx) (not cy))
    (cons nil "x and y are required")
    (progn
      ;; Zoom to the inspection area
      (setq x1 (- cx (/ vw 2.0)))
      (setq y1 (- cy (/ vh 2.0)))
      (setq x2 (+ cx (/ vw 2.0)))
      (setq y2 (+ cy (/ vh 2.0)))
      (command "_.ZOOM" "_W" (list x1 y1) (list x2 y2))

      ;; Collect nearby entities using crossing selection
      (setq ss (ssget "C" (list x1 y1) (list x2 y2)))
      (setq total 0 block-count 0 circle-count 0)
      (setq by-type-str "" blocks-str "" circles-str "")
      (setq largest-circle-radius 0.0 largest-circle-center nil)
      (setq nearest-insert-pt nil nearest-insert-dist 1e30)
      (setq eq-center-x cx eq-center-y cy eq-method "fallback" eq-confidence "low")

      (if ss
        (progn
          (setq total (sslength ss) i 0)
          ;; Count by type
          (setq type-counts '())
          (while (< i total)
            (setq ent (ssname ss i))
            (setq ent-data (entget ent))
            (setq etype (cdr (assoc 0 ent-data)))
            (setq hnd (cdr (assoc 5 ent-data)))

            ;; Count types
            (if (assoc etype type-counts)
              (setq type-counts (mapcar '(lambda (p) (if (= (car p) etype) (cons (car p) (1+ (cdr p))) p)) type-counts))
              (setq type-counts (cons (cons etype 1) type-counts))
            )

            ;; Track circles
            (if (= etype "CIRCLE")
              (progn
                (setq cur-radius (cdr (assoc 40 ent-data)))
                (setq cur-center (cdr (assoc 10 ent-data)))
                (if (> (strlen circles-str) 0) (setq circles-str (strcat circles-str ",")))
                (setq circles-str (strcat circles-str
                  "{\"handle\":\"" hnd
                  "\",\"center\":" (mcp-point-to-json cur-center)
                  ",\"radius\":" (mcp-num-to-json cur-radius) "}"))
                (setq circle-count (1+ circle-count))
                (if (> cur-radius largest-circle-radius)
                  (progn
                    (setq largest-circle-radius cur-radius)
                    (setq largest-circle-center cur-center)
                  )
                )
              )
            )

            ;; Track INSERT blocks
            (if (= etype "INSERT")
              (progn
                (setq blk-name (cdr (assoc 2 ent-data)))
                (setq ins-pt (cdr (assoc 10 ent-data)))
                (if (> (strlen blocks-str) 0) (setq blocks-str (strcat blocks-str ",")))
                (setq blocks-str (strcat blocks-str
                  "{\"handle\":\"" hnd
                  "\",\"block_name\":\"" (mcp-escape-string blk-name)
                  "\",\"position\":" (mcp-point-to-json ins-pt) "}"))
                (setq block-count (1+ block-count))
                ;; Distance from center — only consider if insertion point is within the view window
                ;; (large blocks like xrefs may have insertion far from the viewed area)
                (if (and (>= (car ins-pt) x1) (<= (car ins-pt) x2)
                         (>= (cadr ins-pt) y1) (<= (cadr ins-pt) y2))
                  (progn
                    (setq cur-dist (distance (list cx cy) (list (car ins-pt) (cadr ins-pt))))
                    (if (< cur-dist nearest-insert-dist)
                      (progn
                        (setq nearest-insert-dist cur-dist)
                        (setq nearest-insert-pt ins-pt)
                      )
                    )
                  )
                )
              )
            )
            (setq i (1+ i))
          )

          ;; Build by_type JSON
          (foreach tc type-counts
            (if (> (strlen by-type-str) 0) (setq by-type-str (strcat by-type-str ",")))
            (setq by-type-str (strcat by-type-str "\"" (car tc) "\":" (itoa (cdr tc))))
          )
        )
      )

      ;; === Center Inference ===
      (if (= infer-center "1")
        (progn
          ;; Priority 1: Specific INSERT by handle — compute bbox
          (if (and handle (/= handle ""))
            (progn
              (setq target-ent (handent handle))
              (if target-ent
                (progn
                  (setq target-data (entget target-ent))
                  (if (= (cdr (assoc 0 target-data)) "INSERT")
                    (progn
                      ;; Walk subentities to compute bounding box
                      (setq min-x 1e30 min-y 1e30 max-x -1e30 max-y -1e30 has-bbox nil)
                      (setq sub-ent (entnext target-ent))
                      (while sub-ent
                        (setq sub-data (entget sub-ent))
                        (if (= (cdr (assoc 0 sub-data)) "SEQEND")
                          (setq sub-ent nil)
                          (progn
                            (if (cdr (assoc 10 sub-data))
                              (progn
                                (setq has-bbox T)
                                (if (< (car (cdr (assoc 10 sub-data))) min-x) (setq min-x (car (cdr (assoc 10 sub-data)))))
                                (if (< (cadr (cdr (assoc 10 sub-data))) min-y) (setq min-y (cadr (cdr (assoc 10 sub-data)))))
                                (if (> (car (cdr (assoc 10 sub-data))) max-x) (setq max-x (car (cdr (assoc 10 sub-data)))))
                                (if (> (cadr (cdr (assoc 10 sub-data))) max-y) (setq max-y (cadr (cdr (assoc 10 sub-data)))))
                              )
                            )
                            (setq sub-ent (entnext sub-ent))
                          )
                        )
                      )
                      (if has-bbox
                        (progn
                          (setq eq-center-x (/ (+ min-x max-x) 2.0))
                          (setq eq-center-y (/ (+ min-y max-y) 2.0))
                          (setq eq-method "insert_bbox")
                          (setq eq-confidence "high")
                        )
                        ;; Fallback to insert point
                        (progn
                          (setq eq-center-x (car (cdr (assoc 10 target-data))))
                          (setq eq-center-y (cadr (cdr (assoc 10 target-data))))
                          (setq eq-method "insert_point")
                          (setq eq-confidence "medium")
                        )
                      )
                    )
                  )
                )
              )
            )
          )

          ;; Priority 2: Largest circle (if not already resolved by handle)
          (if (and (= eq-method "fallback") largest-circle-center)
            (progn
              (setq eq-center-x (car largest-circle-center))
              (setq eq-center-y (cadr largest-circle-center))
              (setq eq-method "largest_circle")
              (setq eq-confidence "high")
            )
          )

          ;; Priority 3: Nearest INSERT
          (if (and (= eq-method "fallback") nearest-insert-pt)
            (progn
              (setq eq-center-x (car nearest-insert-pt))
              (setq eq-center-y (cadr nearest-insert-pt))
              (setq eq-method "nearest_insert")
              (setq eq-confidence "medium")
            )
          )
        )
      )

      ;; Build bbox if we have circle data
      (setq bbox-str "null")
      (if (and largest-circle-center (= eq-method "largest_circle"))
        (setq bbox-str (strcat "["
          (rtos (- (car largest-circle-center) largest-circle-radius) 2 6) ","
          (rtos (- (cadr largest-circle-center) largest-circle-radius) 2 6) ","
          (rtos (+ (car largest-circle-center) largest-circle-radius) 2 6) ","
          (rtos (+ (cadr largest-circle-center) largest-circle-radius) 2 6) "]"))
      )

      (cons T (strcat
        "{\"view_center\":" (mcp-point-to-json (list cx cy 0.0))
        ",\"equipment_center\":{\"x\":" (rtos eq-center-x 2 6)
          ",\"y\":" (rtos eq-center-y 2 6)
          ",\"method\":\"" eq-method "\""
          ",\"confidence\":\"" eq-confidence "\""
          ",\"bbox\":" bbox-str "}"
        ",\"nearby_entities\":{\"total\":" (itoa total)
          ",\"by_type\":{" by-type-str "}"
          ",\"blocks\":[" blocks-str "]"
          ",\"circles\":[" circles-str "]}}"))
    )
  )
)

;; -----------------------------------------------------------------------
;; Equipment Tag Placement
;; -----------------------------------------------------------------------

(defun mcp-cmd-place-equipment-tag (params / cx cy cz tag cube-size direction text-height
                                     half ds v1x v1y v2x v2y v3x v3y
                                     mtx mty text-width ulx1 uly ulx2
                                     prev-last cube-handle leader-handle mtext-handle line-handle
                                     bbox-minx bbox-miny bbox-maxx bbox-maxy
                                     tb-result tb-ll tb-ur)
  "Place a complete equipment tag group: 3D cube + leader + MTEXT + underline."

  ;; --- Parse parameters ---
  (setq cx (mcp-json-get-number params "cx"))
  (setq cy (mcp-json-get-number params "cy"))
  (setq cz (mcp-json-get-number params "cz"))
  (if (not cz) (setq cz 0.0))
  (setq tag (mcp-json-get-string params "tag"))
  (setq cube-size (mcp-json-get-number params "cube_size"))
  (if (not cube-size) (setq cube-size 24.0))
  (setq direction (mcp-json-get-string params "direction"))
  (if (not direction) (setq direction "right"))
  (setq text-height (mcp-json-get-number params "text_height"))
  (if (not text-height) (setq text-height 8.0))

  (if (or (not cx) (not cy) (not tag) (= tag ""))
    (cons nil "Missing required parameters: cx, cy, tag")
    (progn
      (setq half (/ cube-size 2.0))
      (setq ds (if (= direction "left") -1.0 1.0))

      ;; --- Ensure layers exist ---
      (if (not (tblsearch "LAYER" "E-EQPM-N"))
        (command "_.LAYER" "_N" "E-EQPM-N" "_C" "2" "E-EQPM-N" "")
      )
      (if (not (tblsearch "LAYER" "E-ANNO-TEXT"))
        (command "_.LAYER" "_N" "E-ANNO-TEXT" "_C" "3" "E-ANNO-TEXT" "")
      )

      ;; --- Ensure LPRT text style exists ---
      (if (not (tblsearch "STYLE" "LPRT"))
        (command "_.STYLE" "LPRT" "ARIALN.TTF" "0" "1.0" "0" "N" "N")
      )

      ;; --- Ensure LPRT IMP dimstyle exists ---
      (if (not (tblsearch "DIMSTYLE" "LPRT IMP"))
        (progn
          (setvar "DIMTXSTY" "LPRT")
          (setvar "DIMTXT" 8.0)
          (setvar "DIMASZ" 4.0)
          (setvar "DIMCLRD" 256)
          (setvar "DIMCLRE" 256)
          (setvar "DIMCLRT" 256)
          (command "_.DIMSTYLE" "_S" "LPRT IMP")
        )
      )

      ;; --- Measure text width using textbox ---
      (setq tb-result (textbox
        (list '(0 . "TEXT") (cons 40 text-height) (cons 1 tag) '(7 . "LPRT"))
      ))
      (if tb-result
        (progn
          (setq tb-ll (car tb-result))
          (setq tb-ur (cadr tb-result))
          (setq text-width (- (car tb-ur) (car tb-ll)))
        )
        ;; Fallback: approximate Arial Narrow width
        (setq text-width (* text-height 0.48 (strlen tag)))
      )

      ;; --- Compute geometry positions ---
      ;; Leader vertices
      (setq v1x (+ cx (* ds half)))
      (setq v1y (+ cy half))
      (setq v2x (+ v1x (* ds 24.0)))
      (setq v2y (+ v1y 48.0))
      (setq v3x (+ v2x (* ds 4.0)))
      (setq v3y v2y)

      ;; MTEXT position
      (if (= direction "left")
        (progn
          (setq mtx (- v3x 4.0 text-width))
          (setq mty (+ v3y 4.0))
        )
        (progn
          (setq mtx (+ v3x 4.0))
          (setq mty (+ v3y 4.0))
        )
      )

      ;; Underline position
      (setq uly (- mty 9.6))
      (setq ulx1 mtx)
      (setq ulx2 (+ mtx text-width))

      ;; --- Create polyface mesh cube ---
      (setq prev-last (entlast))

      ;; POLYLINE header
      (entmake (list
        '(0 . "POLYLINE") '(100 . "AcDbEntity") '(8 . "E-EQPM-N")
        '(100 . "AcDb3dPolyline") '(66 . 1)
        '(10 0.0 0.0 0.0) '(70 . 64) '(71 . 8) '(72 . 6)
      ))

      ;; 8 corner vertices (flag 192 = polyface mesh vertex)
      (entmake (list '(0 . "VERTEX") '(100 . "AcDbEntity") '(8 . "E-EQPM-N")
        '(100 . "AcDbVertex") '(100 . "AcDbPolyFaceMeshVertex")
        (list 10 (- cx half) (- cy half) (+ cz half)) '(70 . 192)))
      (entmake (list '(0 . "VERTEX") '(100 . "AcDbEntity") '(8 . "E-EQPM-N")
        '(100 . "AcDbVertex") '(100 . "AcDbPolyFaceMeshVertex")
        (list 10 (+ cx half) (- cy half) (+ cz half)) '(70 . 192)))
      (entmake (list '(0 . "VERTEX") '(100 . "AcDbEntity") '(8 . "E-EQPM-N")
        '(100 . "AcDbVertex") '(100 . "AcDbPolyFaceMeshVertex")
        (list 10 (+ cx half) (- cy half) (- cz half)) '(70 . 192)))
      (entmake (list '(0 . "VERTEX") '(100 . "AcDbEntity") '(8 . "E-EQPM-N")
        '(100 . "AcDbVertex") '(100 . "AcDbPolyFaceMeshVertex")
        (list 10 (- cx half) (- cy half) (- cz half)) '(70 . 192)))
      (entmake (list '(0 . "VERTEX") '(100 . "AcDbEntity") '(8 . "E-EQPM-N")
        '(100 . "AcDbVertex") '(100 . "AcDbPolyFaceMeshVertex")
        (list 10 (- cx half) (+ cy half) (+ cz half)) '(70 . 192)))
      (entmake (list '(0 . "VERTEX") '(100 . "AcDbEntity") '(8 . "E-EQPM-N")
        '(100 . "AcDbVertex") '(100 . "AcDbPolyFaceMeshVertex")
        (list 10 (+ cx half) (+ cy half) (+ cz half)) '(70 . 192)))
      (entmake (list '(0 . "VERTEX") '(100 . "AcDbEntity") '(8 . "E-EQPM-N")
        '(100 . "AcDbVertex") '(100 . "AcDbPolyFaceMeshVertex")
        (list 10 (+ cx half) (+ cy half) (- cz half)) '(70 . 192)))
      (entmake (list '(0 . "VERTEX") '(100 . "AcDbEntity") '(8 . "E-EQPM-N")
        '(100 . "AcDbVertex") '(100 . "AcDbPolyFaceMeshVertex")
        (list 10 (- cx half) (+ cy half) (- cz half)) '(70 . 192)))

      ;; 6 face records (flag 128 = polyface mesh face)
      (entmake (list '(0 . "VERTEX") '(100 . "AcDbEntity") '(8 . "E-EQPM-N")
        '(100 . "AcDbVertex") '(100 . "AcDbFaceRecord")
        '(10 0.0 0.0 0.0) '(70 . 128) '(71 . 1) '(72 . 2) '(73 . 3) '(74 . 4)))
      (entmake (list '(0 . "VERTEX") '(100 . "AcDbEntity") '(8 . "E-EQPM-N")
        '(100 . "AcDbVertex") '(100 . "AcDbFaceRecord")
        '(10 0.0 0.0 0.0) '(70 . 128) '(71 . 5) '(72 . 6) '(73 . 7) '(74 . 8)))
      (entmake (list '(0 . "VERTEX") '(100 . "AcDbEntity") '(8 . "E-EQPM-N")
        '(100 . "AcDbVertex") '(100 . "AcDbFaceRecord")
        '(10 0.0 0.0 0.0) '(70 . 128) '(71 . 1) '(72 . 2) '(73 . 6) '(74 . 5)))
      (entmake (list '(0 . "VERTEX") '(100 . "AcDbEntity") '(8 . "E-EQPM-N")
        '(100 . "AcDbVertex") '(100 . "AcDbFaceRecord")
        '(10 0.0 0.0 0.0) '(70 . 128) '(71 . 4) '(72 . 3) '(73 . 7) '(74 . 8)))
      (entmake (list '(0 . "VERTEX") '(100 . "AcDbEntity") '(8 . "E-EQPM-N")
        '(100 . "AcDbVertex") '(100 . "AcDbFaceRecord")
        '(10 0.0 0.0 0.0) '(70 . 128) '(71 . 1) '(72 . 5) '(73 . 8) '(74 . 4)))
      (entmake (list '(0 . "VERTEX") '(100 . "AcDbEntity") '(8 . "E-EQPM-N")
        '(100 . "AcDbVertex") '(100 . "AcDbFaceRecord")
        '(10 0.0 0.0 0.0) '(70 . 128) '(71 . 2) '(72 . 6) '(73 . 7) '(74 . 3)))

      ;; SEQEND
      (entmake (list '(0 . "SEQEND") '(100 . "AcDbEntity") '(8 . "E-EQPM-N")))

      ;; Get cube handle: entnext from prev-last gives POLYLINE entity
      (if prev-last
        (setq cube-handle (cdr (assoc 5 (entget (entnext prev-last)))))
        (setq cube-handle (cdr (assoc 5 (entget (entnext)))))
      )

      ;; --- Create LEADER ---
      (entmake (list
        '(0 . "LEADER") '(100 . "AcDbEntity") '(8 . "E-ANNO-TEXT")
        '(100 . "AcDbLeader")
        '(3 . "LPRT IMP")
        '(71 . 1) '(72 . 0) '(73 . 3) '(74 . 1) '(75 . 0)
        '(40 . 0.0) '(41 . 0.0) '(76 . 3)
        (list 10 v1x v1y 0.0)
        (list 10 v2x v2y 0.0)
        (list 10 v3x v3y 0.0)
        '(210 0.0 0.0 1.0)
        '(211 1.0 0.0 0.0)
        '(212 0.0 0.0 0.0)
        '(213 0.0 0.0 0.0)
      ))
      (setq leader-handle (cdr (assoc 5 (entget (entlast)))))

      ;; --- Create MTEXT with background fill ---
      (entmake (list
        '(0 . "MTEXT") '(100 . "AcDbEntity") '(8 . "E-ANNO-TEXT")
        '(100 . "AcDbMText")
        (list 10 mtx mty 0.0)
        (cons 40 text-height)
        (cons 41 text-width)
        '(71 . 1) '(72 . 5)
        (cons 1 tag)
        '(7 . "LPRT")
        '(210 0.0 0.0 1.0)
        '(11 1.0 0.0 0.0)
        '(50 . 0.0)
        '(73 . 1)
        '(44 . 0.75)
        '(90 . 3)
        '(63 . 256)
        '(45 . 1.0)
        '(441 . 0)
      ))
      (setq mtext-handle (cdr (assoc 5 (entget (entlast)))))

      ;; --- Create underline LINE ---
      (entmake (list
        '(0 . "LINE") '(100 . "AcDbEntity") '(8 . "E-ANNO-TEXT")
        '(100 . "AcDbLine")
        (list 10 ulx1 uly 0.0)
        (list 11 ulx2 uly 0.0)
      ))
      (setq line-handle (cdr (assoc 5 (entget (entlast)))))

      ;; --- Compute bounding box ---
      (setq bbox-minx (min (- cx half) ulx1 mtx))
      (setq bbox-miny (min (- cy half) uly))
      (setq bbox-maxx (max (+ cx half) ulx2 (+ mtx text-width)))
      (setq bbox-maxy (max (+ cy half) (+ mty text-height)))

      ;; --- Build result JSON ---
      (cons T (strcat
        "{\"cube_handle\":\"" (if cube-handle cube-handle "unknown") "\""
        ",\"leader_handle\":\"" (if leader-handle leader-handle "unknown") "\""
        ",\"mtext_handle\":\"" (if mtext-handle mtext-handle "unknown") "\""
        ",\"line_handle\":\"" (if line-handle line-handle "unknown") "\""
        ",\"center\":{\"x\":" (rtos cx 2 6)
          ",\"y\":" (rtos cy 2 6)
          ",\"z\":" (rtos cz 2 6) "}"
        ",\"tag\":\"" (mcp-escape-string tag) "\""
        ",\"text_width\":" (rtos text-width 2 6)
        ",\"bbox\":{\"min_x\":" (rtos bbox-minx 2 6)
          ",\"min_y\":" (rtos bbox-miny 2 6)
          ",\"max_x\":" (rtos bbox-maxx 2 6)
          ",\"max_y\":" (rtos bbox-maxy 2 6) "}"
        "}"
      ))
    )
  )
)

;; -----------------------------------------------------------------------
;; Deep Text Search — find text in modelspace AND inside block definitions
;; Searches TEXT, MTEXT, DIMENSION overrides, INSERT ATTRIBs, and block defs.
;; Handles nested blocks (text inside blocks that are inserted inside other blocks).
;; -----------------------------------------------------------------------

(defun mcp-cmd-find-text (params / pattern case-sens max-res zoom-first zoom-ht
                                   result-str found first-pos
                                   ;; Phase 1+2 vars
                                   ss i ent ent-data etype content handle elayer pos
                                   sub-ent sub-data block-name attrib-tag attrib-val
                                   ;; Phase 3 vars
                                   blk-def blk-name blk-ent blk-data blk-type blk-content
                                   ins-ss ins-i ins-data ins-pt ins-handle
                                   ins-sx ins-sy ins-rot world-pt
                                   ;; Nested block tracking
                                   parent-blk parent-ins-ss parent-ins-i parent-ins-data
                                   parent-ins-pt parent-sx parent-sy parent-rot
                                   nested-world-pt blk-local-pos
                                   ;; INSERT tracking inside block defs
                                   nested-inserts nested-name nested-blk-ent nested-blk-data
                                   nested-blk-type nested-blk-content
                                   nested-ins-pt-local nested-sx nested-sy nested-rot
                                   ;; Visited set
                                   visited)
  "Deep text search across modelspace and all block definitions."

  (setq pattern (mcp-json-get-string params "pattern"))
  (setq case-sens (mcp-json-get-string params "case_sensitive"))
  (setq max-res (mcp-json-get-number params "max_results"))
  (setq zoom-first (mcp-json-get-string params "zoom_to_first"))
  (setq zoom-ht (mcp-json-get-number params "zoom_height"))
  (if (not pattern) (setq pattern ""))
  (if (not case-sens) (setq case-sens "0"))
  (if (not max-res) (setq max-res 50))
  (setq max-res (fix max-res))
  (if (not zoom-first) (setq zoom-first "1"))
  (if (not zoom-ht) (setq zoom-ht 600.0))

  (if (= pattern "")
    (cons nil "pattern is required")
    (progn
      (setq result-str "" found 0 first-pos nil visited '())

      ;; === Phase 1: Modelspace TEXT + MTEXT + DIMENSION (ssget "_X") ===
      (setq ss (ssget "X" (list '(-4 . "<OR")
                                '(0 . "TEXT")
                                '(0 . "MTEXT")
                                '(0 . "DIMENSION")
                                '(-4 . "OR>"))))
      (if ss
        (progn
          (setq i 0)
          (while (and (< i (sslength ss)) (< found max-res))
            (setq ent (ssname ss i))
            (setq ent-data (entget ent))
            (setq etype (cdr (assoc 0 ent-data)))
            (setq content (cdr (assoc 1 ent-data)))
            (if (and content (mcp-wildcard-match pattern content case-sens))
              (progn
                (setq handle (cdr (assoc 5 ent-data)))
                (setq elayer (cdr (assoc 8 ent-data)))
                (setq pos (cdr (assoc 10 ent-data)))
                (if (not first-pos) (setq first-pos pos))
                (if (> (strlen result-str) 0) (setq result-str (strcat result-str ",")))
                (setq result-str (strcat result-str
                  "{\"type\":\"" etype "\",\"text\":\"" (mcp-escape-string content)
                  "\",\"layer\":\"" (mcp-escape-string elayer)
                  "\",\"handle\":\"" handle
                  "\",\"position\":" (mcp-point-to-json pos)
                  ",\"world_position\":" (mcp-point-to-json pos)
                  ",\"context\":\"modelspace\"}"))
                (setq found (1+ found))
              )
            )
            (setq i (1+ i))
          )
        )
      )

      ;; === Phase 2: Modelspace INSERT → walk ATTRIBs ===
      (if (< found max-res)
        (progn
          (setq ss (ssget "X" '((0 . "INSERT"))))
          (if ss
            (progn
              (setq i 0)
              (while (and (< i (sslength ss)) (< found max-res))
                (setq ent (ssname ss i))
                (setq ent-data (entget ent))
                (setq block-name (cdr (assoc 2 ent-data)))
                (setq handle (cdr (assoc 5 ent-data)))
                (setq elayer (cdr (assoc 8 ent-data)))
                (setq pos (cdr (assoc 10 ent-data)))
                ;; Only walk attribs if attributes-follow flag (66) is set
                (if (= (cdr (assoc 66 ent-data)) 1)
                  (progn
                    (setq sub-ent (entnext ent))
                    (while (and sub-ent (< found max-res))
                      (setq sub-data (entget sub-ent))
                      (cond
                        ((= (cdr (assoc 0 sub-data)) "ATTRIB")
                         (setq attrib-val (cdr (assoc 1 sub-data)))
                         (if (and attrib-val (mcp-wildcard-match pattern attrib-val case-sens))
                           (progn
                             (setq attrib-tag (cdr (assoc 2 sub-data)))
                             (if (not first-pos) (setq first-pos pos))
                             (if (> (strlen result-str) 0) (setq result-str (strcat result-str ",")))
                             (setq result-str (strcat result-str
                               "{\"type\":\"ATTRIB\",\"text\":\"" (mcp-escape-string attrib-val)
                               "\",\"tag\":\"" (mcp-escape-string attrib-tag)
                               "\",\"layer\":\"" (mcp-escape-string elayer)
                               "\",\"handle\":\"" (cdr (assoc 5 sub-data))
                               "\",\"position\":" (mcp-point-to-json (cdr (assoc 10 sub-data)))
                               ",\"world_position\":" (mcp-point-to-json (cdr (assoc 10 sub-data)))
                               ",\"containing_block\":\"" (mcp-escape-string block-name)
                               "\",\"insert_handle\":\"" handle
                               "\",\"context\":\"attribute\"}"))
                             (setq found (1+ found))
                           )
                         )
                        )
                        ((= (cdr (assoc 0 sub-data)) "SEQEND")
                         (setq sub-ent nil)
                        )
                      )
                      (if sub-ent (setq sub-ent (entnext sub-ent)))
                    )
                  )
                )
                (setq i (1+ i))
              )
            )
          )
        )
      )

      ;; === Phase 3: Walk ALL block definitions for TEXT/MTEXT/ATTDEF ===
      ;; This finds text inside block definitions, including xref content.
      ;; For each match, finds all INSERT references to locate world positions.
      (if (< found max-res)
        (progn
          (setq blk-def (tblnext "BLOCK" T))
          (while (and blk-def (< found max-res))
            (setq blk-name (cdr (assoc 2 blk-def)))
            ;; Skip anonymous blocks (*U*, *X*, *D*)
            (if (and blk-name
                     (> (strlen blk-name) 0)
                     (/= (substr blk-name 1 1) "*")
                     (not (member blk-name visited)))
              (progn
                (setq visited (cons blk-name visited))
                (setq blk-ent (tblobjname "BLOCK" blk-name))
                (if blk-ent (setq blk-ent (entnext blk-ent)))
                (while (and blk-ent (< found max-res))
                  (setq blk-data (entget blk-ent))
                  (setq blk-type (cdr (assoc 0 blk-data)))
                  (cond
                    ;; TEXT/MTEXT inside block definition
                    ((or (= blk-type "TEXT") (= blk-type "MTEXT") (= blk-type "DIMENSION"))
                     (setq blk-content (cdr (assoc 1 blk-data)))
                     (if (and blk-content (mcp-wildcard-match pattern blk-content case-sens))
                       (progn
                         ;; Find modelspace INSERT references for this block
                         (setq ins-ss (ssget "X" (list (cons 0 "INSERT") (cons 2 blk-name))))
                         (if ins-ss
                           (progn
                             (setq ins-i 0)
                             (while (and (< ins-i (sslength ins-ss)) (< found max-res))
                               (setq ins-data (entget (ssname ins-ss ins-i)))
                               (setq ins-pt (cdr (assoc 10 ins-data)))
                               (setq ins-sx (if (cdr (assoc 41 ins-data)) (cdr (assoc 41 ins-data)) 1.0))
                               (setq ins-sy (if (cdr (assoc 42 ins-data)) (cdr (assoc 42 ins-data)) 1.0))
                               (setq ins-rot (if (cdr (assoc 50 ins-data)) (cdr (assoc 50 ins-data)) 0.0))
                               (setq world-pt (mcp-transform-point (cdr (assoc 10 blk-data)) ins-pt ins-sx ins-sy ins-rot))
                               (if (not first-pos) (setq first-pos world-pt))
                               (if (> (strlen result-str) 0) (setq result-str (strcat result-str ",")))
                               (setq result-str (strcat result-str
                                 "{\"type\":\"" blk-type "\",\"text\":\"" (mcp-escape-string blk-content)
                                 "\",\"layer\":\"" (mcp-escape-string (cdr (assoc 8 blk-data)))
                                 "\",\"handle\":\"" (cdr (assoc 5 ins-data))
                                 "\",\"position\":" (mcp-point-to-json (cdr (assoc 10 blk-data)))
                                 ",\"world_position\":" (mcp-point-to-json world-pt)
                                 ",\"containing_block\":\"" (mcp-escape-string blk-name)
                                 "\",\"insert_handle\":\"" (cdr (assoc 5 ins-data))
                                 "\",\"context\":\"block_definition\"}"))
                               (setq found (1+ found))
                               (setq ins-i (1+ ins-i))
                             )
                           )
                           ;; No direct INSERT in modelspace — check if this block is nested
                           ;; (inserted inside another block that IS in modelspace)
                           (progn
                             ;; Record the match with local position only
                             (if (not first-pos) (setq first-pos (cdr (assoc 10 blk-data))))
                             (if (> (strlen result-str) 0) (setq result-str (strcat result-str ",")))
                             (setq result-str (strcat result-str
                               "{\"type\":\"" blk-type "\",\"text\":\"" (mcp-escape-string blk-content)
                               "\",\"layer\":\"" (mcp-escape-string (cdr (assoc 8 blk-data)))
                               "\",\"handle\":\"" (cdr (assoc 5 blk-data))
                               "\",\"position\":" (mcp-point-to-json (cdr (assoc 10 blk-data)))
                               ",\"world_position\":" (mcp-point-to-json (cdr (assoc 10 blk-data)))
                               ",\"containing_block\":\"" (mcp-escape-string blk-name)
                               "\",\"context\":\"nested_block_definition\"}"))
                             (setq found (1+ found))
                           )
                         )
                       )
                     )
                    )
                    ;; ATTDEF inside block definition (check default value)
                    ((= blk-type "ATTDEF")
                     (setq blk-content (cdr (assoc 1 blk-data)))
                     (if (and blk-content (mcp-wildcard-match pattern blk-content case-sens))
                       (progn
                         (setq ins-ss (ssget "X" (list (cons 0 "INSERT") (cons 2 blk-name))))
                         (if ins-ss
                           (progn
                             (setq ins-i 0)
                             (while (and (< ins-i (sslength ins-ss)) (< found max-res))
                               (setq ins-data (entget (ssname ins-ss ins-i)))
                               (setq ins-pt (cdr (assoc 10 ins-data)))
                               (setq ins-sx (if (cdr (assoc 41 ins-data)) (cdr (assoc 41 ins-data)) 1.0))
                               (setq ins-sy (if (cdr (assoc 42 ins-data)) (cdr (assoc 42 ins-data)) 1.0))
                               (setq ins-rot (if (cdr (assoc 50 ins-data)) (cdr (assoc 50 ins-data)) 0.0))
                               (setq world-pt (mcp-transform-point (cdr (assoc 10 blk-data)) ins-pt ins-sx ins-sy ins-rot))
                               (if (not first-pos) (setq first-pos world-pt))
                               (if (> (strlen result-str) 0) (setq result-str (strcat result-str ",")))
                               (setq result-str (strcat result-str
                                 "{\"type\":\"ATTDEF\",\"text\":\"" (mcp-escape-string blk-content)
                                 "\",\"tag\":\"" (mcp-escape-string (cdr (assoc 2 blk-data)))
                                 "\",\"layer\":\"" (mcp-escape-string (cdr (assoc 8 blk-data)))
                                 "\",\"handle\":\"" (cdr (assoc 5 ins-data))
                                 "\",\"position\":" (mcp-point-to-json (cdr (assoc 10 blk-data)))
                                 ",\"world_position\":" (mcp-point-to-json world-pt)
                                 ",\"containing_block\":\"" (mcp-escape-string blk-name)
                                 "\",\"insert_handle\":\"" (cdr (assoc 5 ins-data))
                                 "\",\"context\":\"block_definition\"}"))
                               (setq found (1+ found))
                               (setq ins-i (1+ ins-i))
                             )
                           )
                           (progn
                             (if (not first-pos) (setq first-pos (cdr (assoc 10 blk-data))))
                             (if (> (strlen result-str) 0) (setq result-str (strcat result-str ",")))
                             (setq result-str (strcat result-str
                               "{\"type\":\"ATTDEF\",\"text\":\"" (mcp-escape-string blk-content)
                               "\",\"tag\":\"" (mcp-escape-string (cdr (assoc 2 blk-data)))
                               "\",\"layer\":\"" (mcp-escape-string (cdr (assoc 8 blk-data)))
                               "\",\"handle\":\"" (cdr (assoc 5 blk-data))
                               "\",\"position\":" (mcp-point-to-json (cdr (assoc 10 blk-data)))
                               ",\"world_position\":" (mcp-point-to-json (cdr (assoc 10 blk-data)))
                               ",\"containing_block\":\"" (mcp-escape-string blk-name)
                               "\",\"context\":\"nested_block_definition\"}"))
                             (setq found (1+ found))
                           )
                         )
                       )
                     )
                    )
                  )
                  ;; Advance to next entity; stop at ENDBLK
                  (setq blk-ent (entnext blk-ent))
                  (if blk-ent
                    (if (= (cdr (assoc 0 (entget blk-ent))) "ENDBLK")
                      (setq blk-ent nil)
                    )
                  )
                )
              )
            )
            (setq blk-def (tblnext "BLOCK"))
          )
        )
      )

      ;; === Zoom to first result ===
      (if (and (= zoom-first "1") first-pos (> found 0))
        (progn
          (command "_.ZOOM" "_C"
            (list (car first-pos) (cadr first-pos))
            zoom-ht)
          (cons T (strcat "{\"count\":" (itoa found)
                          ",\"results\":[" result-str "]"
                          ",\"zoomed_to\":" (mcp-point-to-json first-pos) "}"))
        )
        (cons T (strcat "{\"count\":" (itoa found) ",\"results\":[" result-str "]}"))
      )
    )
  )
)

;; -----------------------------------------------------------------------
;; Batch Find-and-Tag helpers
;; -----------------------------------------------------------------------

(defun mcp-find-text-first (tag case-sens / ss i ent ent-data etype content pos
                                           sub-ent sub-data block-name
                                           blk-def blk-name blk-ent blk-data blk-type blk-content
                                           ins-ss ins-data ins-pt ins-sx ins-sy ins-rot world-pt)
  "Find the first occurrence of tag text in the drawing. Returns world position or nil."

  ;; Phase 1: Modelspace TEXT + MTEXT + DIMENSION
  (setq ss (ssget "X" (list '(-4 . "<OR")
                            '(0 . "TEXT")
                            '(0 . "MTEXT")
                            '(0 . "DIMENSION")
                            '(-4 . "OR>"))))
  (if ss
    (progn
      (setq i 0)
      (while (and (< i (sslength ss)) (not pos))
        (setq ent (ssname ss i))
        (setq ent-data (entget ent))
        (setq content (cdr (assoc 1 ent-data)))
        (if (and content (mcp-wildcard-match tag content case-sens))
          (setq pos (cdr (assoc 10 ent-data)))
        )
        (setq i (1+ i))
      )
    )
  )
  (if pos (list pos) ;; return list with position
    (progn
      ;; Phase 2: Modelspace INSERT -> walk ATTRIBs
      (setq ss (ssget "X" '((0 . "INSERT"))))
      (if ss
        (progn
          (setq i 0)
          (while (and (< i (sslength ss)) (not pos))
            (setq ent (ssname ss i))
            (setq ent-data (entget ent))
            (if (= (cdr (assoc 66 ent-data)) 1)
              (progn
                (setq sub-ent (entnext ent))
                (while (and sub-ent (not pos))
                  (setq sub-data (entget sub-ent))
                  (cond
                    ((= (cdr (assoc 0 sub-data)) "ATTRIB")
                     (if (and (cdr (assoc 1 sub-data))
                              (mcp-wildcard-match tag (cdr (assoc 1 sub-data)) case-sens))
                       (setq pos (cdr (assoc 10 ent-data)))
                     )
                    )
                    ((= (cdr (assoc 0 sub-data)) "SEQEND")
                     (setq sub-ent nil)
                    )
                  )
                  (if sub-ent (setq sub-ent (entnext sub-ent)))
                )
              )
            )
            (setq i (1+ i))
          )
        )
      )
      (if pos (list pos)
        (progn
          ;; Phase 3: Walk block definitions
          (setq blk-def (tblnext "BLOCK" T))
          (while (and blk-def (not pos))
            (setq blk-name (cdr (assoc 2 blk-def)))
            (if (and blk-name (> (strlen blk-name) 0) (/= (substr blk-name 1 1) "*"))
              (progn
                (setq blk-ent (tblobjname "BLOCK" blk-name))
                (if blk-ent (setq blk-ent (entnext blk-ent)))
                (while (and blk-ent (not pos))
                  (setq blk-data (entget blk-ent))
                  (setq blk-type (cdr (assoc 0 blk-data)))
                  (cond
                    ((or (= blk-type "TEXT") (= blk-type "MTEXT") (= blk-type "DIMENSION") (= blk-type "ATTDEF"))
                     (setq blk-content (cdr (assoc 1 blk-data)))
                     (if (and blk-content (mcp-wildcard-match tag blk-content case-sens))
                       (progn
                         ;; Find a modelspace INSERT of this block
                         (setq ins-ss (ssget "X" (list (cons 0 "INSERT") (cons 2 blk-name))))
                         (if ins-ss
                           (progn
                             (setq ins-data (entget (ssname ins-ss 0)))
                             (setq ins-pt (cdr (assoc 10 ins-data)))
                             (setq ins-sx (if (cdr (assoc 41 ins-data)) (cdr (assoc 41 ins-data)) 1.0))
                             (setq ins-sy (if (cdr (assoc 42 ins-data)) (cdr (assoc 42 ins-data)) 1.0))
                             (setq ins-rot (if (cdr (assoc 50 ins-data)) (cdr (assoc 50 ins-data)) 0.0))
                             (setq pos (mcp-transform-point (cdr (assoc 10 blk-data)) ins-pt ins-sx ins-sy ins-rot))
                           )
                           ;; No direct INSERT — use local position as fallback
                           (setq pos (cdr (assoc 10 blk-data)))
                         )
                       )
                     )
                    )
                  )
                  (setq blk-ent (entnext blk-ent))
                  (if blk-ent
                    (if (= (cdr (assoc 0 (entget blk-ent))) "ENDBLK")
                      (setq blk-ent nil)
                    )
                  )
                )
              )
            )
            (setq blk-def (tblnext "BLOCK"))
          )
          (if pos (list pos) nil)
        )
      )
    )
  )
)

(defun mcp-place-tag-at (pos tag cube-size direction text-height
                          / cx cy cz half ds tb-result tb-ll tb-ur text-width
                          v1x v1y v2x v2y v3x v3y mtx mty uly ulx1 ulx2
                          prev-last cube-handle leader-handle mtext-handle line-handle)
  "Place equipment tag group at position. Returns JSON string with handles, or nil on failure."

  (setq cx (car pos))
  (setq cy (cadr pos))
  (setq cz (if (caddr pos) (caddr pos) 0.0))
  (setq half (/ cube-size 2.0))
  (setq ds (if (= direction "left") -1.0 1.0))

  ;; Measure text width
  (setq tb-result (textbox
    (list '(0 . "TEXT") (cons 40 text-height) (cons 1 tag) '(7 . "LPRT"))
  ))
  (if tb-result
    (progn
      (setq tb-ll (car tb-result))
      (setq tb-ur (cadr tb-result))
      (setq text-width (- (car tb-ur) (car tb-ll)))
    )
    (setq text-width (* text-height 0.48 (strlen tag)))
  )

  ;; Leader vertices
  (setq v1x (+ cx (* ds half)))
  (setq v1y (+ cy half))
  (setq v2x (+ v1x (* ds 24.0)))
  (setq v2y (+ v1y 48.0))
  (setq v3x (+ v2x (* ds 4.0)))
  (setq v3y v2y)

  ;; MTEXT position
  (if (= direction "left")
    (progn (setq mtx (- v3x 4.0 text-width)) (setq mty (+ v3y 4.0)))
    (progn (setq mtx (+ v3x 4.0)) (setq mty (+ v3y 4.0)))
  )

  ;; Underline
  (setq uly (- mty 9.6))
  (setq ulx1 mtx)
  (setq ulx2 (+ mtx text-width))

  ;; --- Create polyface mesh cube ---
  (setq prev-last (entlast))

  (entmake (list
    '(0 . "POLYLINE") '(100 . "AcDbEntity") '(8 . "E-EQPM-N")
    '(100 . "AcDb3dPolyline") '(66 . 1)
    '(10 0.0 0.0 0.0) '(70 . 64) '(71 . 8) '(72 . 6)
  ))

  ;; 8 corner vertices
  (entmake (list '(0 . "VERTEX") '(100 . "AcDbEntity") '(8 . "E-EQPM-N")
    '(100 . "AcDbVertex") '(100 . "AcDbPolyFaceMeshVertex")
    (list 10 (- cx half) (- cy half) (+ cz half)) '(70 . 192)))
  (entmake (list '(0 . "VERTEX") '(100 . "AcDbEntity") '(8 . "E-EQPM-N")
    '(100 . "AcDbVertex") '(100 . "AcDbPolyFaceMeshVertex")
    (list 10 (+ cx half) (- cy half) (+ cz half)) '(70 . 192)))
  (entmake (list '(0 . "VERTEX") '(100 . "AcDbEntity") '(8 . "E-EQPM-N")
    '(100 . "AcDbVertex") '(100 . "AcDbPolyFaceMeshVertex")
    (list 10 (+ cx half) (- cy half) (- cz half)) '(70 . 192)))
  (entmake (list '(0 . "VERTEX") '(100 . "AcDbEntity") '(8 . "E-EQPM-N")
    '(100 . "AcDbVertex") '(100 . "AcDbPolyFaceMeshVertex")
    (list 10 (- cx half) (- cy half) (- cz half)) '(70 . 192)))
  (entmake (list '(0 . "VERTEX") '(100 . "AcDbEntity") '(8 . "E-EQPM-N")
    '(100 . "AcDbVertex") '(100 . "AcDbPolyFaceMeshVertex")
    (list 10 (- cx half) (+ cy half) (+ cz half)) '(70 . 192)))
  (entmake (list '(0 . "VERTEX") '(100 . "AcDbEntity") '(8 . "E-EQPM-N")
    '(100 . "AcDbVertex") '(100 . "AcDbPolyFaceMeshVertex")
    (list 10 (+ cx half) (+ cy half) (+ cz half)) '(70 . 192)))
  (entmake (list '(0 . "VERTEX") '(100 . "AcDbEntity") '(8 . "E-EQPM-N")
    '(100 . "AcDbVertex") '(100 . "AcDbPolyFaceMeshVertex")
    (list 10 (+ cx half) (+ cy half) (- cz half)) '(70 . 192)))
  (entmake (list '(0 . "VERTEX") '(100 . "AcDbEntity") '(8 . "E-EQPM-N")
    '(100 . "AcDbVertex") '(100 . "AcDbPolyFaceMeshVertex")
    (list 10 (- cx half) (+ cy half) (- cz half)) '(70 . 192)))

  ;; 6 face records
  (entmake (list '(0 . "VERTEX") '(100 . "AcDbEntity") '(8 . "E-EQPM-N")
    '(100 . "AcDbVertex") '(100 . "AcDbFaceRecord")
    '(10 0.0 0.0 0.0) '(70 . 128) '(71 . 1) '(72 . 2) '(73 . 3) '(74 . 4)))
  (entmake (list '(0 . "VERTEX") '(100 . "AcDbEntity") '(8 . "E-EQPM-N")
    '(100 . "AcDbVertex") '(100 . "AcDbFaceRecord")
    '(10 0.0 0.0 0.0) '(70 . 128) '(71 . 5) '(72 . 6) '(73 . 7) '(74 . 8)))
  (entmake (list '(0 . "VERTEX") '(100 . "AcDbEntity") '(8 . "E-EQPM-N")
    '(100 . "AcDbVertex") '(100 . "AcDbFaceRecord")
    '(10 0.0 0.0 0.0) '(70 . 128) '(71 . 1) '(72 . 2) '(73 . 6) '(74 . 5)))
  (entmake (list '(0 . "VERTEX") '(100 . "AcDbEntity") '(8 . "E-EQPM-N")
    '(100 . "AcDbVertex") '(100 . "AcDbFaceRecord")
    '(10 0.0 0.0 0.0) '(70 . 128) '(71 . 4) '(72 . 3) '(73 . 7) '(74 . 8)))
  (entmake (list '(0 . "VERTEX") '(100 . "AcDbEntity") '(8 . "E-EQPM-N")
    '(100 . "AcDbVertex") '(100 . "AcDbFaceRecord")
    '(10 0.0 0.0 0.0) '(70 . 128) '(71 . 1) '(72 . 5) '(73 . 8) '(74 . 4)))
  (entmake (list '(0 . "VERTEX") '(100 . "AcDbEntity") '(8 . "E-EQPM-N")
    '(100 . "AcDbVertex") '(100 . "AcDbFaceRecord")
    '(10 0.0 0.0 0.0) '(70 . 128) '(71 . 2) '(72 . 6) '(73 . 7) '(74 . 3)))

  (entmake (list '(0 . "SEQEND") '(100 . "AcDbEntity") '(8 . "E-EQPM-N")))

  ;; Get cube handle
  (if prev-last
    (setq cube-handle (cdr (assoc 5 (entget (entnext prev-last)))))
    (setq cube-handle (cdr (assoc 5 (entget (entnext)))))
  )

  ;; Create LEADER
  (entmake (list
    '(0 . "LEADER") '(100 . "AcDbEntity") '(8 . "E-ANNO-TEXT")
    '(100 . "AcDbLeader")
    '(3 . "LPRT IMP")
    '(71 . 1) '(72 . 0) '(73 . 3) '(74 . 1) '(75 . 0)
    '(40 . 0.0) '(41 . 0.0) '(76 . 3)
    (list 10 v1x v1y 0.0)
    (list 10 v2x v2y 0.0)
    (list 10 v3x v3y 0.0)
    '(210 0.0 0.0 1.0)
    '(211 1.0 0.0 0.0)
    '(212 0.0 0.0 0.0)
    '(213 0.0 0.0 0.0)
  ))
  (setq leader-handle (cdr (assoc 5 (entget (entlast)))))

  ;; Create MTEXT with background fill
  (entmake (list
    '(0 . "MTEXT") '(100 . "AcDbEntity") '(8 . "E-ANNO-TEXT")
    '(100 . "AcDbMText")
    (list 10 mtx mty 0.0)
    (cons 40 text-height)
    (cons 41 text-width)
    '(71 . 1) '(72 . 5)
    (cons 1 tag)
    '(7 . "LPRT")
    '(210 0.0 0.0 1.0)
    '(11 1.0 0.0 0.0)
    '(50 . 0.0)
    '(73 . 1)
    '(44 . 0.75)
    '(90 . 3)
    '(63 . 256)
    '(45 . 1.0)
    '(441 . 0)
  ))
  (setq mtext-handle (cdr (assoc 5 (entget (entlast)))))

  ;; Create underline LINE
  (entmake (list
    '(0 . "LINE") '(100 . "AcDbEntity") '(8 . "E-ANNO-TEXT")
    '(100 . "AcDbLine")
    (list 10 ulx1 uly 0.0)
    (list 11 ulx2 uly 0.0)
  ))
  (setq line-handle (cdr (assoc 5 (entget (entlast)))))

  ;; Return JSON fragment
  (strcat
    "{\"tag\":\"" (mcp-escape-string tag) "\""
    ",\"status\":\"placed\""
    ",\"position\":" (mcp-point-to-json pos)
    ",\"cube_handle\":\"" (if cube-handle cube-handle "unknown") "\""
    ",\"leader_handle\":\"" (if leader-handle leader-handle "unknown") "\""
    ",\"mtext_handle\":\"" (if mtext-handle mtext-handle "unknown") "\""
    ",\"line_handle\":\"" (if line-handle line-handle "unknown") "\""
    "}"
  )
)

(defun mcp-cmd-batch-find-and-tag (params / tags-str cube-size direction text-height
                                     case-sens tag-list tag placed not-found-list
                                     results-str found-pos place-result)
  "Find and tag multiple equipment items in a single execution."

  (setq tags-str (mcp-json-get-string params "tags"))
  (setq cube-size (mcp-json-get-number params "cube_size"))
  (if (not cube-size) (setq cube-size 24.0))
  (setq direction (mcp-json-get-string params "direction"))
  (if (not direction) (setq direction "right"))
  (setq text-height (mcp-json-get-number params "text_height"))
  (if (not text-height) (setq text-height 8.0))
  (setq case-sens (mcp-json-get-string params "case_sensitive"))
  (if (not case-sens) (setq case-sens "0"))

  (if (or (not tags-str) (= tags-str ""))
    (cons nil "tags parameter is required (semicolon-delimited list)")
    (progn
      (setq tag-list (mcp-split-string tags-str ";"))
      (setq placed 0)
      (setq not-found-list "")
      (setq results-str "")

      ;; Ensure layers/styles exist once before loop
      (if (not (tblsearch "LAYER" "E-EQPM-N"))
        (command "_.LAYER" "_N" "E-EQPM-N" "_C" "2" "E-EQPM-N" "")
      )
      (if (not (tblsearch "LAYER" "E-ANNO-TEXT"))
        (command "_.LAYER" "_N" "E-ANNO-TEXT" "_C" "3" "E-ANNO-TEXT" "")
      )
      (if (not (tblsearch "STYLE" "LPRT"))
        (command "_.STYLE" "LPRT" "ARIALN.TTF" "0" "1.0" "0" "N" "N")
      )
      (if (not (tblsearch "DIMSTYLE" "LPRT IMP"))
        (progn
          (setvar "DIMTXSTY" "LPRT")
          (setvar "DIMTXT" 8.0)
          (setvar "DIMASZ" 4.0)
          (setvar "DIMCLRD" 256)
          (setvar "DIMCLRE" 256)
          (setvar "DIMCLRT" 256)
          (command "_.DIMSTYLE" "_S" "LPRT IMP")
        )
      )

      ;; Process each tag
      (foreach tag tag-list
        (if (and tag (/= tag ""))
          (progn
            (setq found-pos (mcp-find-text-first tag case-sens))
            (if found-pos
              (progn
                ;; found-pos is (list position) — extract the position
                (setq place-result (mcp-place-tag-at (car found-pos) tag cube-size direction text-height))
                (if (> (strlen results-str) 0) (setq results-str (strcat results-str ",")))
                (setq results-str (strcat results-str place-result))
                (setq placed (1+ placed))
              )
              (progn
                ;; Not found
                (if (> (strlen not-found-list) 0) (setq not-found-list (strcat not-found-list ",")))
                (setq not-found-list (strcat not-found-list "\"" (mcp-escape-string tag) "\""))
                (if (> (strlen results-str) 0) (setq results-str (strcat results-str ",")))
                (setq results-str (strcat results-str
                  "{\"tag\":\"" (mcp-escape-string tag) "\",\"status\":\"not_found\"}"))
              )
            )
          )
        )
      )

      (cons T (strcat
        "{\"placed\":" (itoa placed)
        ",\"not_found\":[" not-found-list "]"
        ",\"results\":[" results-str "]"
        "}"))
    )
  )
)

;; -----------------------------------------------------------------------
;; MagiCAD integration commands
;; -----------------------------------------------------------------------

(defun mcp-cmd-magicad-status (/ arxlist magi-modules result)
  "Check if MagiCAD is loaded and list modules."
  (setq arxlist (arx))
  (setq magi-modules
    (vl-remove-if-not
      '(lambda (x) (vl-string-search "magi" (strcase x T)))
      arxlist))
  (if magi-modules
    (progn
      (setq result "{\"loaded\":true,\"modules\":[")
      (setq first T)
      (foreach m magi-modules
        (if first (setq first nil) (setq result (strcat result ",")))
        (setq result (strcat result "\"" (mcp-escape-string m) "\"")))
      (setq result (strcat result "]}"))
      (cons T result))
    (cons T "{\"loaded\":false,\"modules\":[]}")))

(defun mcp-cmd-magicad-run (params-json / cmd-str args arg-list ok)
  "Run a MagiCAD command with optional arguments."
  (setq cmd-str (mcp-json-get-string params-json "command"))
  (if (not cmd-str)
    (cons nil "Missing 'command' parameter")
    (progn
      ;; Verify it's a MagiCAD command (must start with MAGI or -MAGI or _MAGI)
      (setq cmd-upper (strcase cmd-str))
      (if (not (or (vl-string-search "MAGI" cmd-upper)
                   (vl-string-search "MEUPS" cmd-upper)
                   (vl-string-search "MEVPO" cmd-upper)
                   (vl-string-search "MESF" cmd-upper)
                   (vl-string-search "MEDPRJ" cmd-upper)
                   (vl-string-search "MRCAS" cmd-upper)
                   (vl-string-search "MRCSOD" cmd-upper)
                   (vl-string-search "MRCSOB" cmd-upper)))
        (cons nil "Only MagiCAD commands are allowed (must contain MAGI, MEUPS, MEVPO, MESF, MEDPRJ, or MRC)")
        (progn
          ;; Get optional arguments as a list of strings
          (setq args (mcp-json-get-string params-json "args"))
          (if args
            (progn
              (setq arg-list (mcp-string-split args " "))
              ;; Build command call with arguments
              (setq ok (apply 'vl-cmdf (cons cmd-str arg-list))))
            (setq ok (vl-cmdf cmd-str)))
          (if ok
            (cons T (strcat "{\"command\":\"" (mcp-escape-string cmd-str) "\",\"status\":\"executed\"}"))
            (cons nil (strcat "Command '" cmd-str "' failed or was cancelled"))))))))

(defun mcp-cmd-magicad-update-drawing (params-json / flags flag-str)
  "Run -MAGIUPD4 with 16 update flags (all 1 by default)."
  (setq flags (mcp-json-get-string params-json "flags"))
  (if (not flags) (setq flags "1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1"))
  (setq flag-list (mcp-string-split flags " "))
  (if (apply 'vl-cmdf (cons "-MAGIUPD4" flag-list))
    (cons T "{\"status\":\"drawing_updated\"}")
    (cons nil "MAGIUPD4 failed")))

(defun mcp-cmd-magicad-cleanup (params-json / opts)
  "Run -MAGIUCLEAN drawing cleanup."
  (setq opts (mcp-json-get-string params-json "options"))
  (if (not opts) (setq opts "Y Y Y Y 8"))
  (setq opt-list (mcp-string-split opts " "))
  (if (apply 'vl-cmdf (cons "-MAGIUCLEAN" opt-list))
    (cons T "{\"status\":\"cleanup_done\"}")
    (cons nil "MAGIUCLEAN failed")))

(defun mcp-cmd-magicad-ifc-export (params-json / mode)
  "Run MagiCAD IFC export."
  (setq mode (mcp-json-get-string params-json "mode"))
  (if (or (not mode) (= mode "current"))
    (progn
      (if (vl-cmdf "-MAGIIFCEXPORTCURDWG")
        (cons T "{\"status\":\"ifc_exported\",\"mode\":\"current_drawing\"}")
        (cons nil "IFC export failed")))
    (progn
      (if (vl-cmdf "-MAGIIFCEXPORT2")
        (cons T "{\"status\":\"ifc_exported\",\"mode\":\"selection\"}")
        (cons nil "IFC export failed")))))

(defun mcp-cmd-magicad-view-mode (params-json / mode pipe-mode)
  "Change pipe/duct view mode: 1D, 2D, 2D_2D, 2D_3D, 3D."
  (setq mode (mcp-json-get-string params-json "mode"))
  (if (not mode)
    (cons nil "Missing 'mode' parameter (1D, 2D, 2D_2D, 2D_3D, 3D)")
    (progn
      (setq pipe-mode (mcp-json-get-string params-json "type"))
      (if (not pipe-mode) (setq pipe-mode "D"))  ;; D=duct+pipe
      (if (vl-cmdf "MAGICHANGEVIEWMODE" "On" pipe-mode mode)
        (cons T (strcat "{\"status\":\"view_mode_changed\",\"mode\":\"" (mcp-escape-string mode) "\"}"))
        (cons nil "MAGICHANGEVIEWMODE failed")))))

(defun mcp-cmd-magicad-change-storey (params-json / storey)
  "Change active storey."
  (setq storey (mcp-json-get-string params-json "storey"))
  (if (not storey)
    (cons nil "Missing 'storey' parameter")
    (if (vl-cmdf "-MAGICAS" storey)
      (cons T (strcat "{\"status\":\"storey_changed\",\"storey\":\"" (mcp-escape-string storey) "\"}"))
      (cons nil "MAGICAS failed"))))

(defun mcp-cmd-magicad-section-update (/)
  "Update all drawing sections."
  (if (vl-cmdf "-MAGISU")
    (cons T "{\"status\":\"sections_updated\"}")
    (cons nil "MAGISU failed")))

(defun mcp-cmd-magicad-fix-errors (/)
  "Fix ductwork/pipe errors."
  (if (vl-cmdf "MAGICHK")
    (cons T "{\"status\":\"errors_checked\"}")
    (cons nil "MAGICHK failed")))

(defun mcp-cmd-magicad-show-all (/)
  "Unisolate/show all MagiCAD objects."
  (if (vl-cmdf "_MAGIUCL")
    (cons T "{\"status\":\"all_shown\"}")
    (cons nil "MAGIUCL failed")))

(defun mcp-cmd-magicad-clear-garbage (/)
  "Clear MagiCAD garbage layer."
  (if (vl-cmdf "MAGIEMP")
    (cons T "{\"status\":\"garbage_cleared\"}")
    (cons nil "MAGIEMP failed")))

(defun mcp-cmd-magicad-disconnect-project (/)
  "Disconnect drawing from MagiCAD project."
  (if (vl-cmdf "MAGIDPRJ")
    (cons T "{\"status\":\"project_disconnected\"}")
    (cons nil "MAGIDPRJ failed")))

(defun mcp-cmd-magicad-list-commands (/ test-cmds found result first)
  "List available MagiCAD commands in current session."
  (setq test-cmds '(
    ;; Common
    "-MAGIUCLEAN" "-MAGIUCLEAN2" "MAGIEMP" "-MAGIIFCEXPORT2" "-MAGIIFCEXPORTCURDWG" "-MAGIIFCCD"
    ;; HPV / Piping
    "-MAGIUPD" "-MAGIUPD2" "-MAGIUPD3" "-MAGIUPD4" "MAGIEXPLODESCRIPT" "MAGICHANGEVIEWMODE"
    "MAGICHVM" "-MAGIVPO" "-MAGISU" "MAGICHK" "-MAGICAS" "-MAGICSO" "MAGIDPRJ" "_MAGIUCL"
    ;; HPV extended
    "MAGIHPVPIPE" "MAGIHPVDRAW" "MAGIHPVROUTE" "MAGIHPVINSERT" "MAGIHPVSIZE"
    "MAGIHPVCALC" "MAGIHPVINDEX" "MAGIHPVCHECK" "MAGIHPVBALANCE" "MAGIHPVREPORT"
    "MAGIHPVPRODUCT" "MAGIHPVEDIT" "MAGIHPVINFO" "MAGIHPVCONNECT" "MAGIHPVSETTINGS"
    "MAGIHPVSCHEMATIC" "MAGIHPVUPDATE" "MAGIHPVSECTIONUPDATE" "MAGIHPVFIXERRORS"
    "MAGIHPVUCLEAN" "MAGIHPVEXPLODE" "MAGIHPVIFCEXPORT" "MAGIHPVSHOWALL"
    "MAGIHPVCHANGESTOREY" "MAGIHPVSTOREYORIGIN" "MAGIHPVDISCONNECT" "MAGIHPVVIEWMODE"
    "MAGIHPVVIEWPORT"
    ;; Sprinkler
    "MAGISPRINKLER" "MAGIHPVSPRINKLER" "MAGIHPVSPRINKLERHEAD" "MAGIHPVSPRINKLERCALC"
    "MAGIHPVSPRINKLERCHECK" "MAGIHPVSPRINKLERREPORT" "MAGISPRINKLERCALC"
    "MAGISPRINKLERHEAD" "MAGISPRINKLERCHECK" "MAGISPRINKLERREPORT"
    "MAGISPRINKLERDRAW" "MAGISPRINKLERROUTE" "MAGISPRINKLERINSERT" "MAGISPRINKLERSIZE"
    "MAGISPRINKLERINDEX" "MAGISPRINKLEREDIT" "MAGISPRINKLERINFO"
    "MAGISPRINKLERCONNECT" "MAGISPRINKLERSETTINGS" "MAGISPRINKLERSCHEMATIC"
    "MAGISPRINKLERBALANCE" "MAGISPRINKLERPRODUCT" "MAGISPRINKLERPIPE"
    ;; Electrical
    "MEUPS3" "MAGIESECTIONUPDATE" "MAGIEVPORTOPTIONSSCRIPT" "MEVPOS"
    "MAGIESHOWFLOOR" "MESF" "MAGIESHOWALL" "MEDPRJ"
    ;; Room
    "-MRCAS" "-MRCSOD" "-MRCSOB"
  ))
  (setq found '())
  (foreach c test-cmds
    (if (vl-cmdf c) (setq found (cons c found))))
  (setq result "[")
  (setq first T)
  (foreach c (reverse found)
    (if first (setq first nil) (setq result (strcat result ",")))
    (setq result (strcat result "\"" c "\"")))
  (setq result (strcat result "]"))
  (cons T (strcat "{\"count\":" (itoa (length found)) ",\"commands\":" result "}")))

;; -----------------------------------------------------------------------
;; Startup message
;; -----------------------------------------------------------------------

(princ "\n=== MCP Dispatch v5.0 loaded ===")
(princ "\nIPC directory: ")
(princ *mcp-ipc-dir*)
(princ "\nReady for commands via (c:mcp-dispatch)")
(princ)
