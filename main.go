package main

import (
	"bufio"
	"bytes"
	"encoding/json"
	"fmt"
	"image"
	_ "image/gif"
	_ "image/jpeg"
	_ "image/png"
	"mime"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"strconv"
	"strings"
	"time"

	chroma "github.com/alecthomas/chroma/v2"
	"github.com/alecthomas/chroma/v2/lexers"
	"github.com/alecthomas/chroma/v2/styles"
	docx "github.com/fumiama/go-docx"
	gitignore "github.com/sabhiram/go-gitignore"
	"github.com/sergi/go-diff/diffmatchpatch"
)

// ─── Colour helpers ──────────────────────────────────────────────────────────

func printGreen(text string) { fmt.Printf("\033[92m%s\033[0m\n", text) }
func printYellow(text string) { fmt.Printf("\033[93m%s\033[0m\n", text) }
func printRed(text string) { fmt.Printf("\033[91m%s\033[0m\n", text) }

// ─── Config ──────────────────────────────────────────────────────────────────

type Config struct {
	Language             string `json:"language"`
	Verbose              bool   `json:"verbose"`
	DiffFont             string `json:"diff_font"`
	DiffFontSize         int    `json:"diff_font_size"`
	OpenAfterCreation    bool   `json:"open_after_creation"`
	HeadingLevel         int    `json:"heading_level"`
	AddColor             string `json:"add_color"`
	RemoveColor          string `json:"remove_color"`
	NeutralColor         string `json:"neutral_color"`
	AddSymbol            string `json:"add_symbol"`
	RemoveSymbol         string `json:"remove_symbol"`
	NeutralSymbol        string `json:"neutral_symbol"`
	FileEncoding         string `json:"file_encoding"`
	IncludeFirstCommit   bool   `json:"include_first_commit"`
	GDDIgnoreFileName    string `json:"gdd_ignore_file_name"`
	IncludeUnchangedLines bool  `json:"include_unchanged_lines"`
	IncludeImages        bool   `json:"include_images"`
	InsertPageBreaks     bool   `json:"insert_page_breaks"`
	PygmentsStyle        string `json:"pygments_style"`
}

// configDefaults returns a Config struct pre-filled with the Python defaults.
// json.Unmarshal will override any field that is explicitly set in the JSON.
func configDefaults() Config {
	return Config{
		Language:              "en",
		Verbose:               false,
		DiffFont:              "Courier New",
		DiffFontSize:          8,
		OpenAfterCreation:     false,
		HeadingLevel:          2,
		AddColor:              "D0FFD0",
		RemoveColor:           "FFD0D0",
		NeutralColor:          "F5F5F5",
		AddSymbol:             "+",
		RemoveSymbol:          "-",
		NeutralSymbol:         "=",
		FileEncoding:          "utf-8",
		IncludeFirstCommit:    false,
		GDDIgnoreFileName:     ".gddignore",
		IncludeUnchangedLines: true,
		IncludeImages:         true,
		InsertPageBreaks:      true,
		PygmentsStyle:         "default",
	}
}

// ─── Lang ────────────────────────────────────────────────────────────────────

type Lang struct {
	BinaryFileSkipped    string `json:"binary_file_skipped"`
	Diffs                string `json:"diffs"`
	EnterCommit1         string `json:"enter_commit1"`
	EnterCommit2         string `json:"enter_commit2"`
	ErrorInsertingImage  string `json:"error_inserting_image"`
	EnterOutputDocx      string `json:"enter_output_docx"`
	EnterTargetDir       string `json:"enter_target_dir"`
	ErrorOpeningFile     string `json:"error_opening_file"`
	ErrorRemovingFile    string `json:"error_removing_file"`
	ErrorSavingFile      string `json:"error_saving_file"`
	Exiting              string `json:"exiting"`
	File                 string `json:"file"`
	GitChangesReport     string `json:"git_changes_report"`
	ImageChanged         string `json:"image_changed"`
	InvalidTargetDir     string `json:"invalid_target_dir"`
	Legend               string `json:"legend"`
	LegendAdd            string `json:"legend_add"`
	LegendNeutral        string `json:"legend_neutral"`
	LegendRemove         string `json:"legend_remove"`
	No                   string `json:"no"`
	NoChangesFound       string `json:"no_changes_found"`
	NoGitRepoFound       string `json:"no_git_repo_found"`
	NoSignificantChanges string `json:"no_significant_changes"`
	OutputExists         string `json:"output_exists"`
	PressEnterToRetry    string `json:"press_enter_to_retry"`
	ProcessingDone       string `json:"processing_done"`
	ProcessingFile       string `json:"processing_file"`
	ReportGeneratedOn    string `json:"report_generated_on"`
	SavingReport         string `json:"saving_report"`
	StillContinue        string `json:"still_continue"`
	Title                string `json:"title"`
	UsingDefaultOutput   string `json:"using_default_output"`
	UsingFirstCommit     string `json:"using_first_commit"`
	UsingLastCommit      string `json:"using_last_commit"`
	Yes                  string `json:"yes"`
}

// fmtStr replaces {key} placeholders in a template string.
func fmtStr(tmpl string, kv ...string) string {
	for i := 0; i+1 < len(kv); i += 2 {
		tmpl = strings.ReplaceAll(tmpl, "{"+kv[i]+"}", kv[i+1])
	}
	return tmpl
}

// ─── Input helpers ───────────────────────────────────────────────────────────

var stdinReader = bufio.NewReader(os.Stdin)

func readLine() string {
	line, _ := stdinReader.ReadString('\n')
	return strings.TrimRight(line, "\r\n")
}

func promptLine(prompt string) string {
	fmt.Print(prompt + " ")
	return readLine()
}

func askYesNo(prompt string, lang Lang) bool {
	for {
		answer := strings.ToLower(strings.TrimSpace(promptLine(prompt)))
		if answer == lang.Yes {
			return true
		}
		if answer == lang.No {
			return false
		}
	}
}

// ─── Binary / image detection ────────────────────────────────────────────────

// isBinaryBytes mirrors Python's is_binary_string.
func isBinaryBytes(data []byte) bool {
	// text chars: {7,8,9,10,12,13,27} | range(0x20,0x100) - {0x7f}
	for _, b := range data {
		ok := b == 7 || b == 8 || b == 9 || b == 10 || b == 12 || b == 13 || b == 27 ||
			(b >= 0x20 && b != 0x7f)
		if !ok {
			return true
		}
	}
	return false
}

// isImageFile mirrors Python's is_image_file.
func isImageFile(filename string) bool {
	ext := strings.ToLower(filepath.Ext(filename))
	if ext == "" {
		return false
	}
	mimeType := mime.TypeByExtension(ext)
	return strings.HasPrefix(mimeType, "image/")
}

// ─── Git helpers ─────────────────────────────────────────────────────────────

func gitRun(args ...string) string {
	out, err := exec.Command("git", args...).Output()
	if err != nil {
		return ""
	}
	return strings.TrimSpace(string(out))
}

func gitBytes(args ...string) []byte {
	out, _ := exec.Command("git", args...).Output()
	return out
}

// ─── Diff ────────────────────────────────────────────────────────────────────

// splitLines splits text into lines, stripping carriage returns, matching
// Python's str.splitlines() behaviour for \r\n and \n line endings.
func splitLines(text string) []string {
	text = strings.ReplaceAll(text, "\r\n", "\n")
	text = strings.ReplaceAll(text, "\r", "\n")
	if text == "" {
		return nil
	}
	lines := strings.Split(text, "\n")
	// splitlines() in Python does NOT include a trailing empty element when
	// the text ends with a newline, so we mirror that.
	if len(lines) > 0 && lines[len(lines)-1] == "" {
		lines = lines[:len(lines)-1]
	}
	return lines
}

type diffLine struct {
	prefix string // " ", "+", or "-"
	text   string
}

// buildDiffLines builds a slice of diff lines (prefix + content), mirroring
// Python's SequenceMatcher-based approach.
func buildDiffLines(oldContent, newContent []string, includeUnchanged bool) []diffLine {
	// Join into full text for DiffLinesToChars
	oldText := strings.Join(oldContent, "\n")
	newText := strings.Join(newContent, "\n")
	if len(oldContent) > 0 {
		oldText += "\n"
	}
	if len(newContent) > 0 {
		newText += "\n"
	}

	dmp := diffmatchpatch.New()
	a, b, lineArray := dmp.DiffLinesToChars(oldText, newText)
	diffs := dmp.DiffMain(a, b, false)
	diffs = dmp.DiffCharsToLines(diffs, lineArray)

	var result []diffLine

	for _, d := range diffs {
		lines := splitLines(d.Text)
		switch d.Type {
		case diffmatchpatch.DiffEqual:
			if !includeUnchanged {
				continue
			}
			for _, l := range lines {
				result = append(result, diffLine{" ", l})
			}
		case diffmatchpatch.DiffDelete:
			for _, l := range lines {
				result = append(result, diffLine{"-", l})
			}
		case diffmatchpatch.DiffInsert:
			for _, l := range lines {
				result = append(result, diffLine{"+", l})
			}
		}
	}
	return result
}

// ─── Document helpers ────────────────────────────────────────────────────────

// headingStyle returns the OOXML style name for a heading level.
func headingStyle(level int) string {
	if level < 1 {
		level = 1
	}
	if level > 9 {
		level = 9
	}
	return "Heading" + strconv.Itoa(level)
}

// halfPtStr converts a point size to the OOXML half-point string.
func halfPtStr(pt int) string {
	return strconv.Itoa(pt * 2)
}

// nilBorder returns a WTableBorder with val="nil" (removes the border).
func nilBorder() *docx.WTableBorder {
	return &docx.WTableBorder{Val: "nil"}
}

// setCellBorders sets specific borders on a table cell to "nil".
func setCellBorders(cell *docx.WTableCell, top, left, bottom, right bool) {
	borders := &docx.WTableBorders{}
	if top {
		borders.Top = nilBorder()
	}
	if left {
		borders.Left = nilBorder()
	}
	if bottom {
		borders.Bottom = nilBorder()
	}
	if right {
		borders.Right = nilBorder()
	}
	cell.TableCellProperties.TableBorders = borders
}

// addSingleRun adds a single text run directly with xml:space="preserve".
// This bypasses AddText's split-on-newline / skip-empty-string logic.
func addSingleRun(p *docx.Paragraph, text string) *docx.Run {
	t := &docx.Text{Text: text, XMLSpace: "preserve"}
	run := &docx.Run{
		RunProperties: &docx.RunProperties{},
		Children:      []interface{}{t},
	}
	p.Children = append(p.Children, run)
	return run
}

// applyRunStyle applies font name, size, color, bold, italic to a run.
func applyRunStyle(run *docx.Run, font string, sizePt int, entry chroma.StyleEntry) {
	run.Font(font, font, font, "")
	run.Size(halfPtStr(sizePt))
	run.SizeCs(halfPtStr(sizePt))

	if entry.Bold == chroma.Yes {
		run.Bold()
	}
	if entry.Italic == chroma.Yes {
		run.Italic()
	}
	if entry.Colour.IsSet() {
		hex := entry.Colour.String() // "#rrggbb"
		run.Color(hex[1:])           // strip leading "#"
	}
}

// ─── Legend table ────────────────────────────────────────────────────────────

func addLegendTable(f *docx.Docx, cfg Config, lang Lang) {
	type legendRow struct {
		label  string
		color  string
		symbol string
	}
	rows := []legendRow{
		{lang.LegendAdd, cfg.AddColor, cfg.AddSymbol},
		{lang.LegendRemove, cfg.RemoveColor, cfg.RemoveSymbol},
	}
	if cfg.IncludeUnchangedLines {
		rows = append(rows, legendRow{lang.LegendNeutral, cfg.NeutralColor, cfg.NeutralSymbol})
	}

	tbl := f.AddTable(len(rows), 2, 0, nil)

	for i, row := range rows {
		cells := tbl.TableRows[i].TableCells

		// Label cell
		if len(cells[0].Paragraphs) == 0 {
			cells[0].AddParagraph()
		}
		p0 := cells[0].Paragraphs[0]
		addSingleRun(p0, row.label)

		// Symbol cell with background color
		cells[1].Shade("clear", "auto", row.color)
		if len(cells[1].Paragraphs) == 0 {
			cells[1].AddParagraph()
		}
		p1 := cells[1].Paragraphs[0]
		r := addSingleRun(p1, row.symbol)
		r.Font(cfg.DiffFont, cfg.DiffFont, cfg.DiffFont, "")
		r.Size(halfPtStr(cfg.DiffFontSize))
		r.SizeCs(halfPtStr(cfg.DiffFontSize))
	}
}

// ─── Diff table ──────────────────────────────────────────────────────────────

// A4 usable width in twips with 1-inch margins on an A4 page (11906 twips wide).
// symbol column ≈ 0.57 cm = 323 twips; code column gets the rest.
const (
	symbolColTwips = 323
	pageWidthTwips = 9026  // A4 width (11906) minus 2×1-inch margins (2×1440)
	codeColTwips   = pageWidthTwips - symbolColTwips
)

func addDiffTable(
	f *docx.Docx,
	diffLines []diffLine,
	cfg Config,
	style *chroma.Style,
	lexer chroma.Lexer,
) {
	if len(diffLines) == 0 {
		return
	}

	// Filter unchanged lines again (inside the table builder) to match Python.
	if !cfg.IncludeUnchangedLines {
		var filtered []diffLine
		for _, dl := range diffLines {
			if dl.prefix != " " {
				filtered = append(filtered, dl)
			}
		}
		diffLines = filtered
	}

	total := len(diffLines)
	if total == 0 {
		return
	}

	rowHeights := make([]int64, total)  // all 0 = auto height
	colWidths := []int64{symbolColTwips, codeColTwips}
	tbl := f.AddTableTwips(rowHeights, colWidths, 0, nil)

	for idx, dl := range diffLines {
		var fill, symbol string
		switch dl.prefix {
		case "+":
			fill = cfg.AddColor
			symbol = cfg.AddSymbol
		case "-":
			fill = cfg.RemoveColor
			symbol = cfg.RemoveSymbol
		default:
			fill = cfg.NeutralColor
			symbol = cfg.NeutralSymbol
		}

		row := tbl.TableRows[idx]
		symCell := row.TableCells[0]
		codeCell := row.TableCells[1]

		// Apply background shading to both cells.
		symCell.Shade("clear", "auto", fill)
		codeCell.Shade("clear", "auto", fill)

		// Border removal: remove the divider between symbol and code cells, and
		// remove internal horizontal borders between consecutive same-colour rows.
		symTop := idx != 0
		symBottom := idx != total-1
		codeTop := idx != 0
		codeBottom := idx != total-1
		setCellBorders(symCell, symTop, false, symBottom, true)
		setCellBorders(codeCell, codeTop, true, codeBottom, false)

		// Symbol paragraph.
		if len(symCell.Paragraphs) == 0 {
			symCell.AddParagraph()
		}
		symPara := symCell.Paragraphs[0]
		symRun := addSingleRun(symPara, symbol)
		symRun.Font(cfg.DiffFont, cfg.DiffFont, cfg.DiffFont, "")
		symRun.Size(halfPtStr(cfg.DiffFontSize))
		symRun.SizeCs(halfPtStr(cfg.DiffFontSize))

		// Code paragraph with syntax highlighting.
		if len(codeCell.Paragraphs) == 0 {
			codeCell.AddParagraph()
		}
		codePara := codeCell.Paragraphs[0]
		setParagraphSpacing(codePara)

		codeContent := dl.text // no prefix character (the prefix is in dl.prefix)

		// Replace empty code content with a non-breaking space (matches Python).
		if codeContent == "" {
			codeContent = "\u00A0"
		}

		// Tokenise and apply syntax highlighting.
		tokeniseAndAdd(codePara, codeContent, cfg, style, lexer)
	}
}

// tokeniseAndAdd tokenises codeContent and adds styled runs to para.
func tokeniseAndAdd(
	para *docx.Paragraph,
	codeContent string,
	cfg Config,
	style *chroma.Style,
	lexer chroma.Lexer,
) {
	it, err := lexer.Tokenise(nil, codeContent)
	if err != nil {
		// Fallback: add plain run.
		r := addSingleRun(para, codeContent)
		r.Font(cfg.DiffFont, cfg.DiffFont, cfg.DiffFont, "")
		r.Size(halfPtStr(cfg.DiffFontSize))
		r.SizeCs(halfPtStr(cfg.DiffFontSize))
		return
	}

	for tok := it(); tok != chroma.EOF; tok = it() {
		value := strings.TrimRight(tok.Value, "\n")
		if value == "" {
			value = "\u00A0" // non-breaking space, mirrors Python
		}

		entry := style.Get(tok.Type)
		r := addSingleRun(para, value)
		applyRunStyle(r, cfg.DiffFont, cfg.DiffFontSize, entry)
	}
}

// ─── Image helper ────────────────────────────────────────────────────────────

func addImage(f *docx.Docx, data []byte, imageName string, lang Lang) {
	// Verify the bytes are a decodable image.
	_, _, err := image.DecodeConfig(bytes.NewReader(data))
	if err != nil {
		p := f.AddParagraph()
		r := addSingleRun(p, fmtStr(lang.ErrorInsertingImage, "image_name", imageName))
		r.Italic()
		return
	}

	p := f.AddParagraph()
	if _, err = p.AddInlineDrawing(data); err != nil {
		pErr := f.AddParagraph()
		r := addSingleRun(pErr, fmtStr(lang.ErrorInsertingImage, "image_name", imageName))
		r.Italic()
	}
	// Centre the image paragraph (alignment 1 = center in python-docx).
	p.Justification("center")
}

// ─── Paragraph spacing helpers ───────────────────────────────────────────────

// setParagraphSpacing zeroes out spacing before the paragraph (mirrors
// python-docx's paragraph_format.space_before = Pt(0)).
// The fumiama/go-docx Spacing struct does not expose w:after, so we only set
// w:before here.
func setParagraphSpacing(p *docx.Paragraph) {
	if p.Properties == nil {
		p.Properties = &docx.ParagraphProperties{}
	}
	if p.Properties.Spacing == nil {
		p.Properties.Spacing = &docx.Spacing{}
	}
	p.Properties.Spacing.Before = 0
}

// ─── Main ────────────────────────────────────────────────────────────────────

func main() {
	// Locate script directory (executable path).
	scriptDir, _ := filepath.Abs(filepath.Dir(os.Args[0]))

	// ── Load config ──────────────────────────────────────────────────────
	configFile := filepath.Join(scriptDir, "config.json")
	if _, err := os.Stat(configFile); os.IsNotExist(err) {
		printRed("Error: Configuration file not found: " + configFile)
		os.Exit(1)
	}
	cfgData, err := os.ReadFile(configFile)
	if err != nil {
		printRed("Error reading config: " + err.Error())
		os.Exit(1)
	}
	cfg := configDefaults()
	if err := json.Unmarshal(cfgData, &cfg); err != nil {
		printRed("Error parsing config: " + err.Error())
		os.Exit(1)
	}
	// Zero values for int/string fields that should have a non-zero default.
	if cfg.HeadingLevel == 0 {
		cfg.HeadingLevel = 2
	}
	if cfg.DiffFontSize == 0 {
		cfg.DiffFontSize = 8
	}
	if cfg.DiffFont == "" {
		cfg.DiffFont = "Courier New"
	}
	if cfg.AddColor == "" {
		cfg.AddColor = "D0FFD0"
	}
	if cfg.RemoveColor == "" {
		cfg.RemoveColor = "FFD0D0"
	}
	if cfg.NeutralColor == "" {
		cfg.NeutralColor = "F5F5F5"
	}
	if cfg.AddSymbol == "" {
		cfg.AddSymbol = "+"
	}
	if cfg.RemoveSymbol == "" {
		cfg.RemoveSymbol = "-"
	}
	if cfg.NeutralSymbol == "" {
		cfg.NeutralSymbol = "="
	}
	if cfg.FileEncoding == "" {
		cfg.FileEncoding = "utf-8"
	}
	if cfg.GDDIgnoreFileName == "" {
		cfg.GDDIgnoreFileName = ".gddignore"
	}
	if cfg.PygmentsStyle == "" {
		cfg.PygmentsStyle = "default"
	}

	// ── Load language file ───────────────────────────────────────────────
	langDir := filepath.Join(scriptDir, "lang")
	if _, err := os.Stat(langDir); os.IsNotExist(err) {
		printRed("Error: Language directory not found: " + langDir)
		os.Exit(1)
	}
	langChoice := cfg.Language
	if langChoice == "" {
		langChoice = "en"
	}
	langFile := filepath.Join(langDir, langChoice+".json")
	if _, err := os.Stat(langFile); os.IsNotExist(err) {
		printRed("Language '" + langChoice + "' not found, defaulting to English if available.")
		langFile = filepath.Join(langDir, "en.json")
	}
	langData, err := os.ReadFile(langFile)
	if err != nil {
		printRed("Error reading language file: " + err.Error())
		os.Exit(1)
	}
	var lang Lang
	if err := json.Unmarshal(langData, &lang); err != nil {
		printRed("Error parsing language file: " + err.Error())
		os.Exit(1)
	}

	// ── Pygments/chroma style ────────────────────────────────────────────
	// Map "default" (pygments name) to "pygments" (chroma equivalent).
	chromaStyleName := cfg.PygmentsStyle
	if chromaStyleName == "default" {
		chromaStyleName = "pygments"
	}
	chromaStyle := styles.Get(chromaStyleName)

	// ── Banner ───────────────────────────────────────────────────────────
	printGreen(lang.Title)

	// ── Target directory ─────────────────────────────────────────────────
	var targetDir string
	for {
		targetDir = strings.TrimSpace(promptLine(lang.EnterTargetDir))
		info, err := os.Stat(targetDir)
		if err != nil || !info.IsDir() {
			printRed(lang.InvalidTargetDir)
			continue
		}
		// Check for .git
		gitDir := filepath.Join(targetDir, ".git")
		if _, err := os.Stat(gitDir); os.IsNotExist(err) {
			printRed(fmtStr(lang.NoGitRepoFound, "target_dir", targetDir))
			if askYesNo(lang.StillContinue, lang) {
				break
			}
			continue
		}
		if err := os.Chdir(targetDir); err != nil {
			printRed("Error changing directory: " + err.Error())
			continue
		}
		break
	}

	// ── GDDIgnore ─────────────────────────────────────────────────────────
	gddIgnorePath := filepath.Join(targetDir, cfg.GDDIgnoreFileName)
	var ignorer *gitignore.GitIgnore
	if _, err := os.Stat(gddIgnorePath); err == nil {
		ignorer, _ = gitignore.CompileIgnoreFile(gddIgnorePath)
	}

	// ── First commit ─────────────────────────────────────────────────────
	commit1Input := strings.TrimSpace(promptLine(lang.EnterCommit1))
	commit1Specified := commit1Input != ""
	commit1 := commit1Input

	if !commit1Specified {
		commit1 = gitRun("rev-list", "--max-parents=0", "HEAD")
	}

	// Determine if commit1 is the very first commit.
	veryFirstCommit := gitRun("rev-list", "--max-parents=0", "HEAD")
	isVeryFirst := (commit1 == veryFirstCommit)

	if !strings.HasSuffix(commit1, "^") && !strings.Contains(commit1, "~") && commit1 != "HEAD" {
		if isVeryFirst && cfg.IncludeFirstCommit {
			commit1 = "4b825dc642cb6eb9a060e54bf8d69288fbee4904" // empty tree
		} else if cfg.IncludeFirstCommit {
			commit1 = commit1 + "^"
		}
	}

	if !commit1Specified {
		fmt.Println(fmtStr(lang.UsingFirstCommit, "commit1", commit1))
	}

	// ── Last commit ───────────────────────────────────────────────────────
	commit2Input := strings.TrimSpace(promptLine(lang.EnterCommit2))
	commit2 := commit2Input
	if commit2 == "" {
		commit2 = gitRun("rev-parse", "HEAD")
		fmt.Println(fmtStr(lang.UsingLastCommit, "commit2", commit2))
	}

	// ── Output file ───────────────────────────────────────────────────────
	outputDocxInput := strings.TrimSpace(promptLine(lang.EnterOutputDocx))
	outputDocx := outputDocxInput
	if outputDocx == "" {
		outputDocx = filepath.Join(scriptDir, "output.docx")
		fmt.Println(fmtStr(lang.UsingDefaultOutput, "output_docx", outputDocx))
	}

	// ── Changed files ─────────────────────────────────────────────────────
	changedFilesRaw := gitRun("diff", "--name-only", commit1, commit2)
	var changedFiles []string
	if changedFilesRaw != "" {
		for _, f := range strings.Split(changedFilesRaw, "\n") {
			f = strings.TrimSpace(f)
			if f == "" || f == cfg.GDDIgnoreFileName {
				continue
			}
			if ignorer != nil && ignorer.MatchesPath(f) {
				continue
			}
			changedFiles = append(changedFiles, f)
		}
	}

	if len(changedFiles) == 0 {
		printYellow(fmtStr(lang.NoChangesFound, "commit1", commit1, "commit2", commit2))
		os.Exit(0)
	}

	// ── Document setup ────────────────────────────────────────────────────
	doc := docx.New().WithDefaultTheme()

	// Title heading (level 1).
	titlePara := doc.AddParagraph()
	titlePara.Style(headingStyle(1))
	addSingleRun(titlePara,
		lang.GitChangesReport+" ("+commit1+" → "+commit2+")")

	// Report date paragraph.
	datePara := doc.AddParagraph()
	addSingleRun(datePara,
		fmtStr(lang.ReportGeneratedOn, "date", time.Now().Format("2006-01-02 15:04:05")))

	// Legend heading.
	legendHeading := doc.AddParagraph()
	legendHeading.Style(headingStyle(cfg.HeadingLevel))
	addSingleRun(legendHeading, lang.Legend)
	addLegendTable(doc, cfg, lang)

	// Page break before diffs.
	doc.AddParagraph().AddPageBreaks()

	// Diffs heading.
	diffsHeading := doc.AddParagraph()
	diffsHeading.Style(headingStyle(cfg.HeadingLevel))
	addSingleRun(diffsHeading, lang.Diffs)

	// ── Overwrite check ───────────────────────────────────────────────────
	if _, err := os.Stat(outputDocx); err == nil {
		if !askYesNo(fmtStr(lang.OutputExists, "output_docx", outputDocx), lang) {
			fmt.Println(lang.Exiting)
			os.Exit(0)
		}
		// Check the file is writable (retry loop mirrors Python).
		for {
			f, err := os.OpenFile(outputDocx, os.O_WRONLY, 0)
			if err == nil {
				f.Close()
				break
			}
			printRed(fmtStr(lang.ErrorRemovingFile,
				"output_docx", outputDocx, "error", err.Error()))
			promptLine(lang.PressEnterToRetry)
		}
	}

	// ── Per-file loop ─────────────────────────────────────────────────────
	for index, file := range changedFiles {
		if index != 0 && cfg.InsertPageBreaks {
			doc.AddParagraph().AddPageBreaks()
		}

		fileHeading := doc.AddParagraph()
		fileHeading.Style(headingStyle(cfg.HeadingLevel + 1))
		addSingleRun(fileHeading, lang.File+": "+file)

		if cfg.Verbose {
			fmt.Println(fmtStr(lang.ProcessingFile, "file", file))
		}

		// Fetch binary content from git.
		oldBytes := gitBytes("show", commit1+":"+file)
		newBytes := gitBytes("show", commit2+":"+file)

		// Binary / image handling.
		if isBinaryBytes(newBytes) || isBinaryBytes(oldBytes) {
			if isImageFile(file) {
				if string(oldBytes) != string(newBytes) {
					p := doc.AddParagraph()
					r := addSingleRun(p, lang.ImageChanged)
					r.Italic()
					if cfg.IncludeImages {
						addImage(doc, newBytes, file, lang)
					}
				}
			} else {
				p := doc.AddParagraph()
				r := addSingleRun(p, lang.BinaryFileSkipped)
				r.Italic()
			}
			if cfg.Verbose {
				printGreen(fmtStr(lang.ProcessingDone, "file", file))
			}
			continue
		}

		// Decode text content.
		oldContent := splitLines(string(oldBytes))
		newContent := splitLines(string(newBytes))

		// Choose lexer: try filename first, then analyse content.
		sample := strings.Join(newContent, "\n")
		if sample == "" {
			sample = strings.Join(oldContent, "\n")
		}
		lexer := lexers.Match(file)
		if lexer == nil {
			lexer = lexers.Analyse(sample)
		}
		if lexer == nil {
			lexer = lexers.Fallback
		}
		lexer = chroma.Coalesce(lexer)

		// Build diff lines.
		diffs := buildDiffLines(oldContent, newContent, cfg.IncludeUnchangedLines)

		if len(diffs) > 0 {
			addDiffTable(doc, diffs, cfg, chromaStyle, lexer)
		} else {
			noChangePara := doc.AddParagraph()
			r := addSingleRun(noChangePara, lang.NoSignificantChanges)
			r.Italic()
		}

		if cfg.Verbose {
			printGreen(fmtStr(lang.ProcessingDone, "file", file))
		}
	}

	// ── Save document ─────────────────────────────────────────────────────
	outFile, err := os.Create(outputDocx)
	if err != nil {
		printRed(fmtStr(lang.ErrorSavingFile, "output_docx", outputDocx, "error", err.Error()))
		os.Exit(1)
	}
	if _, err := doc.WriteTo(outFile); err != nil {
		outFile.Close()
		printRed(fmtStr(lang.ErrorSavingFile, "output_docx", outputDocx, "error", err.Error()))
		os.Exit(1)
	}
	outFile.Close()

	printGreen(fmtStr(lang.SavingReport, "output_docx", outputDocx))

	// ── Open after creation ───────────────────────────────────────────────
	if cfg.OpenAfterCreation {
		var openCmd *exec.Cmd
		switch runtime.GOOS {
		case "windows":
			openCmd = exec.Command("cmd", "/c", "start", "", outputDocx)
		case "darwin":
			openCmd = exec.Command("open", outputDocx)
		default: // linux and others
			openCmd = exec.Command("xdg-open", outputDocx)
		}
		if err := openCmd.Start(); err != nil {
			printRed(fmtStr(lang.ErrorOpeningFile, "file", outputDocx, "error", err.Error()))
		}
	}
}
