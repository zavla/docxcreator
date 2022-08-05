// to debug a service run
// curl.exe -T .\ex1.txt --url http://127.0.0.1:8080/docxcreator --verbose --output .\out2.docx

//

// to create a service
// New-Service -Name docxcreator -BinaryPathName F:\Zavla_VB\go\src\INVENTY\INVENTAR\INVENTAR.exe -Description "creats docx documents" -StartupType Manual

// to run a service example
// .\INVENTAR.exe --PathToTemplates ".\" --bindAddressPort 192.168.3.53:8080 --logfile ".\log53.txt"

package main

import (
	"bytes"
	"encoding/json"
	"errors"
	"flag"
	"fmt"
	"io"
	"io/fs"
	"io/ioutil"
	"log"
	"net/http"
	"os"
	"path/filepath"
	"sort"
	"strconv"
	"strings"

	"baliance.com/gooxml/document"
	"golang.org/x/sys/windows/svc"
	"golang.org/x/sys/windows/svc/eventlog"
)

// input we expect as json
type input struct {
	// name of docx document used as template
	DocxTemplateName string
	// if true returns merge-fields from docx template
	ShowFields bool
	// Header will be MailMerged into new document
	Header map[string]string
	// Table1 will be rows of some table in document
	Table1 []map[int]string
}

type serviceConfig struct {
	pathToTemplates string
	fullPathLogFile string
	bindAddressPort string
}

// responsestruct is returned by this http service as JSON
type responseStruct struct {
	Status  string
	Message string
	Data    []byte
}

// func (r responseStruct) MarshalJSON() ([]byte, error) {
// 	sb := strings.Builder{}
// 	sb.WriteString("{")

// 	sb.WriteString("\"status\":\"")
// 	sb.WriteString(r.Status)
// 	sb.WriteString("\",")

// 	sb.WriteString("\"message\":\"")
// 	sb.WriteString(strings.ReplaceAll(r.Message, "\"", "\\\""))
// 	sb.WriteString("\",")

// 	enc := json.NewEncoder(&sb)
// 	sb.WriteString("\n\"Data\":")
// 	enc.Encode(r.Data)

// 	sb.WriteString("}")
// 	s := sb.String()
// 	logfile.Print(s, "\n")
// 	return []byte(s), nil
// }

var currentConfig serviceConfig
var logfile *log.Logger
var elog *eventlog.Log
var errSomeInfo error = errors.New("info")

const (
	thisServiceName = "docxcreator"

	// a kind of default value for the MergeFields in document that are not set by user input
	constUnderscore = "___________"
)

// make a JSON for the error
func makeresponse(statusString, message string, Data []byte) []byte {

	resp := &responseStruct{
		Status:  statusString,
		Message: message,
		Data:    Data,
	}
	b, err := json.Marshal(&resp)
	if err != nil {
		const op = "makeresponse"
		logfile.Printf("in %s, %s", op, err)
	}
	return b
}

// START
func main() {

	// when starts as a service PathToTemplates should not contain `\"`
	// because \" breaks the flag.Parse()
	fset, pathToTemplates, inputJSON, bindAddressPort, logfilename := defineParameters()

	fset.Parse(os.Args[1:])

	isInteractive, err := svc.IsAnInteractiveSession()
	if err != nil {
		log.Fatalf("%s", err)
	}

	if !isInteractive { //starts as a service by Windows Service Control Mngr

		err := eventlog.InstallAsEventCreate(thisServiceName, eventlog.Info|eventlog.Error)
		if err != nil {
			log.Fatalf("%s\n%s\n", err, "can't InstallAsEventCreate ...")
		}

		elog, err = eventlog.Open(thisServiceName)
		if err != nil {
			log.Fatalf("%s\n%s\n", err, "can't open the event log...")
		}
		defer elog.Close()

		elog.Info(1, "service "+thisServiceName+" is starting...")

		// start as a service requires a log file name
		if err := checkpath(logfilename, true); err != nil {

			bwriter := bytes.NewBuffer(make([]byte, 0, 500))
			fset.SetOutput(bwriter) // a windows service has no terminal to write to
			bwriter.WriteString(err.Error())
			bwriter.WriteRune('\n')
			fset.Usage() // help message to the windows event log
			elog.Info(1, bwriter.String())
			log.Fatalf("%s\n", err.Error())
		}

	}

	// default log to the stdout if its ran as a CLI

	flog, err := setupLogfile(logfilename)
	if err != nil {
		log.Fatalf("can't create log file, %s, %s\n", *logfilename, err)
		if !isInteractive {
			elog.Info(1, "service "+thisServiceName+" could not open its log file: "+(*logfilename))
		}
	}
	defer flog.Close()

	logfile = log.New(flog, "", log.Ldate|log.Ltime)

	if !isInteractive {
		elog.Info(1, "service "+thisServiceName+" started a log file: "+(*logfilename))
	}
	// logfile is ready

	if err := checkpath(pathToTemplates, true); err != nil {
		fset.Usage()
		if !isInteractive {
			elog.Info(1, "service "+thisServiceName+" needs parameter pathToTemplates")
		}
		logfile.Fatalf("parameter PathToTemplates is expected, %s\n", err)
	}

	// global var
	currentConfig = serviceConfig{
		pathToTemplates: *pathToTemplates,
		fullPathLogFile: *logfilename,
		bindAddressPort: *bindAddressPort,
	}

	if *inputJSON == "" { //started as a windows service or as a CLI http server

		if isInteractive { // by user from terminal
			logfile.Println("http server has started")
			err = runHTTP(currentConfig.bindAddressPort)
			logfile.Printf("Http server Exited: %s\n", err)

		} else { // by a windows services manager

			// runs server on other goroutine
			go runHTTP(currentConfig.bindAddressPort)

			// runs SCM responder on other goroutine
			err := svc.Run(thisServiceName, &Tservice{currentConfig: currentConfig})

			logfile.Printf("%s\n", "service "+thisServiceName+" exited.")
			if err != nil {
				logfile.Printf("%s %s\n", "service "+thisServiceName+" exited with error: ", err)
				log.Fatalf("%s", err)
			}
		}

	}
	// here we are only if user starts us as CLI with an input specified as a file
	{
		//command line mode, expecting file as an input

		// open, read, validate
		f, err := os.Open(*inputJSON)
		if err != nil {
			logfile.Println(helpText())
			logfile.Fatalf("%s\n%s\n", "Error: can't open json file with input.", err)
		}
		defer f.Close()

		databytes, err := ioutil.ReadAll(f)
		if err != nil {
			logfile.Fatalf("%s\n%s\n", "Can't read input file.", err)
		}

		inputStru, err := validate_input(databytes)
		if err != nil {
			logfile.Println(err)
			logfile.Fatalln(helpText())
		}

		outputfile := "outfile.docx" //fixed output docx file name
		newfilefullpath := filepath.Join(".\\", outputfile)

		if err := backupAfile(newfilefullpath); err != nil {
			logfile.Println(err)
		}

		outfile, err := os.OpenFile(newfilefullpath, os.O_CREATE|os.O_WRONLY, 0660)
		if err != nil {
			logfile.Fatalf("error: can't output to the file %s, %s", newfilefullpath, err)
			os.Exit(1)
		}
		defer outfile.Close()

		info, err := Createdocx(outfile, inputStru, *pathToTemplates)
		if err != nil {
			logfile.Printf("%s, %s\n", err, string(info))
			os.Exit(1)
		}
		logfile.Printf("OK: %s\n", newfilefullpath)
	}

}

func checkpath(pathToTemplates *string, required bool) error {
	if required && *pathToTemplates == "" {
		return errors.New("empty")
	}
	// it is always required to Abs the user input
	abspathToTemplates, err := filepath.Abs(*pathToTemplates)
	if err != nil {
		return errors.New("path error")
	}
	pathToTemplates = &abspathToTemplates
	return nil
}

func defineParameters() (*flag.FlagSet, *string, *string, *string, *string) {
	fset := flag.NewFlagSet(os.Args[0], flag.ContinueOnError)

	pathToTemplates := fset.String("PathToTemplates", "", "path to templates files. Template name expected in incoming json.")
	jsonFileName := fset.String("input", "", `file with JSON data, utf-8. Used in CLI mode.`)

	bindAddressPort := fset.String("bindAddressPort", "127.0.0.1:8080", "bind service to address and port. Used in service mode.")

	logfilename := fset.String("logfile", "", "path and name to service log file. Used in service mode.")
	return fset, pathToTemplates, jsonFileName, bindAddressPort, logfilename
}

func backupAfile(name string) error {
	if _, err := os.Stat(name); os.IsNotExist(err) {
		return nil
	}
	//rename
	newname := name + ".bak"
	i := 1
	for {
		if _, err := os.Stat(newname); os.IsNotExist(err) {
			break
		}
		//file exists
		newname = name + ".bak" + strconv.Itoa(i)
		i++
	}
	var err error
	if err = os.Rename(name, newname); err != nil {
		err = fmt.Errorf("can't rename %s -> %s, error: %w", name, newname, err)
	}
	return err
}

func handlerhttp(w http.ResponseWriter, r *http.Request) {

	if r.Method != "POST" || !strings.HasPrefix(r.URL.Path, "/docxcreator") {
		w.WriteHeader(http.StatusBadRequest)
		w.Write(makeresponse("error", "A POST method must be used to the /docxcreator endpoint", []byte{}))
		return
	}
	rdr := r.Body
	defer rdr.Close()

	// the action on /docxcreator url writes into w io.Writer by itself or returns info.
	// everything comes in in JSON body.
	info, err := action(w, rdr)

	switch err {
	case errSomeInfo:
		w.WriteHeader(http.StatusAccepted)
		w.Write(makeresponse("OK", info, nil))
		return
	case nil:
		//OK
		//response body was filled with bytes of returned file
		return
	}

	w.WriteHeader(http.StatusBadRequest)
	// err goes to a log
	logfile.Printf("%s", err)
	// info goes to user
	w.Write(makeresponse("error", info, nil))
	return

}

func action(w io.Writer, toreadbytes io.ReadCloser) (infoforUser string, err error) {
	const op = "action"

	var inputBody []byte // data read from request body
	inputBody, err = ioutil.ReadAll(toreadbytes)
	if err != nil {
		var errCode string = "UnableToProcessRequest" //the declaration of errCode must be in this form

		logfile.Printf("errCode=%s, %s", errCode, err)

		return fmt.Sprintf("Server was unable to process your request, errCode=%s", errCode), err
	}

	inputStru, err := validate_input(inputBody)
	if err != nil {
		var errCode string = "InputValidationError"
		// do not log every validation error
		return fmt.Sprintf("Your input JSON doesn't pass validation, errCode=%s, %s, %s", errCode, err, helpText()), err //this error goes to clients
	}

	// Createdocx creates docx files and writes them into w.
	info, err := Createdocx(w, inputStru, currentConfig.pathToTemplates)

	switch {
	case errors.Is(err, errSomeInfo):
		//have some information for the user, not an error
		return string(info), err
	case err != nil:
		logfile.Printf("failed to create %s, %s\n", inputStru.DocxTemplateName, err)
		var errCode string = "CantCreateDocx"
		//user gets an errCode and general message
		return fmt.Sprintf("Server encountered an error while creating a document, errCode=%s, %s", errCode, info), err

	}

	logfile.Printf("successfully created %s\n", inputStru.DocxTemplateName)
	return string(info), nil
}

func getDocumentPtrFromTemplate(templatename, pathToTemplates string) (*document.Document, error) {
	//opens template
	fullPathToTemplate := filepath.Join(pathToTemplates, templatename)
	doc, err := document.Open(fullPathToTemplate)
	if err != nil {
		return nil, err
	}
	return doc, nil
}

func alltemplates() ([]fs.DirEntry, error) {
	dir := currentConfig.pathToTemplates
	osDir := os.DirFS(dir)
	docs, err := fs.ReadDir(osDir, ".")
	if err != nil {
		return nil, err
	}
	return docs, nil

}

func helpText() string {
	sb := strings.Builder{}
	const exampleJson = `{
		"DocxTemplateName": "youtemplatenamehere.docx",
		"ShowFields": false,
		"Header": {
			"Номер": "ЮХ000000084",
			"Дата": "19.08.2021 11:31:20",
			"НомерДоговора": "123",
			"ИтогоСумма": "645,26"
		},
		"Table1": [
			{
				"1": "1",
				"2": "Ремкомплект MM для редуктора СО2 Premium",
				"3": "шт",
				"4": "2",
				"5": "322,63",
				"6": "645,26"
			}
		]
	}`
	sb.WriteString(`Help message. A service expects a JSON input in the following form:
`)

	bb := bytes.NewBuffer(make([]byte, 0, len(exampleJson)))
	if err := json.Compact(bb, []byte(exampleJson)); err != nil {
		logfile.Printf("json.Compact returned error, %s", err)
	}
	sb.Write(bb.Bytes())

	sb.WriteString(`
You must specify DocxTemplateName.
You may use the following DocxTemplateName values:
`)
	docs, err := alltemplates()
	if err != nil {
		sb.WriteString("The list of available document templates was not generated due to internal error.")
	} else {
		for k, _ := range docs {
			sb.WriteString(docs[k].Name())
			sb.WriteRune('\n')
		}
		if len(docs) == 0 {
			sb.WriteString("This service doesn't have any ./templates/*.docx files.")
		}
	}
	return sb.String()
}

func validate_input(databytes []byte) (*input, error) {
	inputStru, err := convertIntoinput(databytes) //converts json to struct
	if err != nil {
		return nil, err
	}

	if inputStru.DocxTemplateName == "" {

		return nil, errors.New("input JSON requires a DocxTemplateName tag")
	}
	return &inputStru, nil
}

// Createdocx creates docx document through gooxml, fills MergeFields with data from input JSON
// and adds rows into tables in docx document. Tables are searched by the row content: text "1" "2" in first
// two cells.
// New rows are filled from struct "input".
func Createdocx(w io.Writer, inputStru *input, pathToTemplates string) (infoUser []byte, err error) {
	const op = "Createdocx"

	doc, err := getDocumentPtrFromTemplate(inputStru.DocxTemplateName, pathToTemplates)
	if err != nil {
		var errCode string = "BadTemplate"
		err2 := fmt.Errorf("%s, template %s, errCode=%s, %w\n", op, inputStru.DocxTemplateName, errCode, err)
		logfile.Println(err2)
		return []byte(fmt.Sprintf("Server was unable to open your template %s", inputStru.DocxTemplateName)),
			err2
	}

	const convertMergeFieldsIntoText = "0123456789101112131415161718192021222324252627282930"
	helpmessage := make([]string, 0, 20)
	// first, merge fields from template documet will be filled with predefined value "__________"
	for _, v := range doc.MergeFields() {

		if inputStru.ShowFields {
			helpmessage = append(helpmessage, v)
		}
		if strings.Contains(convertMergeFieldsIntoText, v) {
			// if user made MergeFields in template with names like 1,2,3 ...
			// replace this merge fields with just text 1,2,3...
			inputStru.Header[v] = v
		}
		if _, has := inputStru.Header[v]; !has {
			// if user doesn't supply a value for the mergefield  will make it more visible
			// with default value for it
			inputStru.Header[v] = constUnderscore
		}
	}
	if inputStru.ShowFields { // client requested help for available MergeFields
		sort.Strings(helpmessage)
		return []byte(strings.Join(helpmessage, ";\n")), errSomeInfo
	}

	doc.MailMerge(inputStru.Header) // inserts values into the document by MailMerge

	if inputStru.Table1 != nil && len(inputStru.Table1) != 0 {
		// WORKING WITH TABLE
		// searches the table by row content "1 2 3 4 5"
		var tabfound bool
		var tabindex int
		var totalcells int      // how many cells are in fact in the table in the document template
		var insrow document.Row //a row where to insert new rows
		var insrowIndex int

		tabfound, tabindex, insrow, totalcells, insrowIndex = findOurTable(doc) // searches the table

		if !tabfound {
			err := errors.New("template file doesn't have a Table object with a row with 1,2,3 values in its cells.")
			logfile.Printf("%s", err)
			return []byte{}, err
		}
		tab := doc.Tables()[tabindex]

		insertnewRows(&tab, insrow, insrowIndex, totalcells, inputStru.Table1)

		// remove the row after which we have inserted new rows
		tabWithRemove := TableWithDelete{&tab}
		tabWithRemove.RemoveRow(insrow)
	}

	// saves new dowcument into io.Writer
	err = doc.Save(w)
	if err != nil {
		logfile.Printf("error: can't write into io.writer, %s\n", err)
		return []byte{}, err
	}
	logfile.Printf("Successuly served %q", inputStru.DocxTemplateName)
	return []byte{}, nil
}

// convertIntoinput unmarshals a json into struct
func convertIntoinput(bstr []byte) (input, error) {
	var v input
	err := json.Unmarshal(bstr, &v)
	if err != nil {
		logfile.Printf("%s\n%s\n", "Error: can't use your json.", err)
		return v, err

	}
	return v, nil

}

// findOurTable seeks the table with the row with cells with text "1 2 3 4 5"
func findOurTable(doc *document.Document) (bool, int, document.Row, int, int) {

	var tabfound bool
	var tabindex int
	var retrow document.Row // a row that we have found
	var totalcells int      // actual number of cells in the row
	var retrowIndex int     //index of the row that we have found

	tables := doc.Tables()
	for i, tab := range tables {

		rows := tab.Rows()
		for ri, row := range rows {
			if len(row.Cells()) < 2 {
				continue
			}
			cell0 := row.Cells()[0]
			cell1 := row.Cells()[1]
			if len(cell0.Paragraphs()) < 1 || len(cell1.Paragraphs()) < 1 {
				continue
			}
			if len(cell0.Paragraphs()[0].Runs()) < 1 || len(cell1.Paragraphs()[0].Runs()) < 1 {
				continue
			}
			col1text := cell0.Paragraphs()[0].Runs()[0].Text()
			col2text := cell1.Paragraphs()[0].Runs()[0].Text()
			if col1text == "1" && col2text == "2" {
				tabfound = true
				tabindex = i
				totalcells = len(row.Cells()) // how many cells may be passed into the row
				retrow = row
				retrowIndex = ri
				goto wayout
			}
		}

	}
wayout:
	return tabfound, tabindex, retrow, totalcells, retrowIndex

}

// I add a new method to an external package type Table
type TableWithDelete struct {
	*document.Table
}

// our additional method RemoveRow
func (t *TableWithDelete) RemoveRow(r document.Row) {
	for i, rc := range t.X().EG_ContentRowContent {

		if len(rc.Tr) > 0 && r.X() == rc.Tr[0] {
			if i+1 < len(t.X().EG_ContentRowContent) {
				copy(t.X().EG_ContentRowContent[i:], t.X().EG_ContentRowContent[i+1:])
			}
			t.X().EG_ContentRowContent = t.X().EG_ContentRowContent[:len(t.X().EG_ContentRowContent)-1]
			break
		}
	}
}

// insertnewRows adds rows into tab from slice or rows (maps)
func insertnewRows(tab *document.Table, startrow document.Row, startRowIndex int, countcells int, newrows []map[int]string) {
	currow := startrow
	lastrow := false
	if len(tab.X().EG_ContentRowContent) == startRowIndex+1 {
		// this is the last row in the table
		lastrow = true
	}
	for _, datamap := range newrows {
		var nrow document.Row
		if lastrow {
			nrow = tab.AddRow() // faster then InsertRowAfter

		} else {
			nrow = tab.InsertRowAfter(currow)
		}
		currow = nrow
		for nc := 1; nc <= countcells; nc++ {

			ncell := nrow.AddCell()
			npar := ncell.AddParagraph()

			nrun := npar.AddRun()
			nrun.AddText(datamap[nc]) // nc is 1,2, a column number passed in incoming json
		}
	}
}

func runHTTP(bindAddressPort string) error {
	var hand http.HandlerFunc = handlerhttp
	//starts serving
	err := http.ListenAndServe(bindAddressPort, hand)
	return err

}

func setupLogfile(logfilename *string) (*os.File, error) {
	flog := os.Stdout
	if *logfilename != "" {
		logfilenamefull, err := filepath.Abs(*logfilename)
		if err != nil {
			return flog, fmt.Errorf("bad log file name, %w\n", err)
		}

		// begin log
		flog, err = os.OpenFile(logfilenamefull, os.O_RDWR|os.O_CREATE, 0660)
		if err != nil {
			return flog, fmt.Errorf("can't create log file, %s, %w\n", logfilenamefull, err)
		}
	}
	return flog, nil
}

// When Word saves a document, it removes all unused styles.  This means to
// copy the styles from an existing document, you must first create a
// document that contains text in each style of interest.
// for _, s := range doc.Styles.Styles() {
// 	fmt.Println("style", s.Name(), "has ID of", s.StyleID(), "type is", s.Type())
// }
