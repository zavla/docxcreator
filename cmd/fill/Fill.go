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

// input we expecting as json
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

// error response struct, returned by service as json
type responseStruct struct {
	Error   string
	Data    []byte
	StrData string
}

var currentConfig serviceConfig
var logfile *log.Logger
var elog *eventlog.Log

const (
	thisServiceName = "docxcreator"
	constUnderscore = "___________"
)

// jsons the error message
func makeresponse(rerr string, StrData string, Data []byte) []byte {
	resp := &responseStruct{
		Error:   rerr,
		StrData: StrData,
		Data:    Data,
	}
	b, err := json.Marshal(&resp)
	if err != nil {
		logfile.Printf("%s", err)
	}
	return b
}

func main() {

	// when starts as a service PathToTemplates should not contain `\"`
	// because \" breaks the flag.Parse()
	fset := flag.NewFlagSet(os.Args[0], flag.ContinueOnError)

	pathToTemplates := fset.String("PathToTemplates", "", "path to templates files. Template name expected in incoming json.")
	jsonFileName := fset.String("input", "", `file with JSON data, utf-8. Used in CLI mode.`)

	bindAddressPort := fset.String("bindAddressPort", "127.0.0.1:8080", "bind service to address and port. Used in service mode.")

	logfilename := fset.String("logfile", "", "path and name to service log file. Used in service mode.")

	fset.Parse(os.Args[1:])

	isInteractive, err := svc.IsAnInteractiveSession()
	if err != nil {
		//log.Fatalf("%s", err)
	}
	if !isInteractive { //starts as a service

		err := eventlog.InstallAsEventCreate(thisServiceName, eventlog.Info|eventlog.Error)
		if err != nil {
			log.Printf("%s\n%s\n", err, "can't InstallAsEventCreate ...")
		}
		elog, err = eventlog.Open(thisServiceName)
		if err != nil {
			log.Printf("%s\n%s\n", err, "started without event log...")
		}
		defer elog.Close()
		elog.Info(1, "service "+thisServiceName+" is starting...")

		if *logfilename == "" {
			errstr := "parameter logfile is empty"

			bwriter := bytes.NewBuffer(make([]byte, 0, 200))
			fset.SetOutput(bwriter)

			fmt.Fprintf(bwriter, "%s\n", errstr)
			fset.Usage()
			elog.Info(1, bwriter.String())
			log.Fatalf("%s\n", errstr)
		}

	}

	logfilenamefull := ""
	flog := os.Stdout
	if *logfilename != "" {
		logfilenamefull, err := filepath.Abs(*logfilename)
		if err != nil {
			log.Fatalf("%s\n", "bad log file name")
		}
		*pathToTemplates, err = filepath.Abs(*pathToTemplates)
		if err != nil {
			log.Fatalf("%s\n", *pathToTemplates, err)
		}

		// begin log
		flog, err = os.OpenFile(logfilenamefull, os.O_RDWR|os.O_CREATE, 0660)
		if err != nil {
			if !isInteractive {
				elog.Info(1, "service "+thisServiceName+" could not open its log file: "+(logfilenamefull))
			}
			log.Fatalf("%s\n%s\n", "Error: can't create log file: "+(logfilenamefull), err)
		}
	}

	logfile = log.New(flog, "", log.Ldate|log.Ltime)
	logfile.Printf("%s", "in main() ...")
	if !isInteractive {
		elog.Info(1, "service "+thisServiceName+" started a log file: "+(logfilenamefull))
	}
	// logfile is ready

	if *pathToTemplates == "" {
		fset.Usage()
		if !isInteractive {
			elog.Info(1, "service "+thisServiceName+" needs parameter pathToTemplates")
		}
		logfile.Fatalf("%s\n", "Error: PathToTemplates - path to a folder with docx templates needed.")
	}

	// global var
	currentConfig = serviceConfig{
		pathToTemplates: *pathToTemplates,
		fullPathLogFile: *logfilename,
		bindAddressPort: *bindAddressPort,
	}

	if *jsonFileName == "" { //starts as a service, uses http

		if isInteractive { // ran by user
			err = runHTTP(currentConfig.bindAddressPort)
			fmt.Printf("%s, %s\n", "Http server Exited:", err)

		} else { // ran by a services manager

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

	} else { //command line mode, expecting file as an input

		// open, read, validate
		f, err := os.Open(*jsonFileName)
		if err != nil {
			logfile.Println(helpText())
			logfile.Fatalf("%s\n%s\n", "Error: can't open json file with input.", err)
		}
		databytes, err := ioutil.ReadAll(f)
		if err != nil {
			logfile.Fatalf("%s\n%s\n", "Can't read input file.", err)
		}
		inputStru, err := validate_input(databytes)
		if err != nil {
			logfile.Fatalln(err)
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

		info, err := CreateDocxFromStruct(outfile, inputStru, *pathToTemplates)
		if err != nil {
			logfile.Printf("%s, %s\n", err, string(info))
			os.Exit(1)
		}
		logfile.Printf("OK: %s\n", newfilefullpath)
	}

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

	}
	var err error
	if err = os.Rename(name, newname); err != nil {
		err = fmt.Errorf("can't rename %s -> %s, error: %w", name, newname, err)
	}
	return err
}

func handlerhttp(w http.ResponseWriter, r *http.Request) {
	//action(w, )
	if r.URL.Path == "/docxcreator" {
		if r.Method == "GET" {
			w.WriteHeader(http.StatusBadRequest)
			w.Write(makeresponse(fmt.Sprintf("%s", "POST should be used."), "", []byte{}))
			return
		}
		rdr := r.Body
		defer rdr.Close()

		// the action on /docxcreator url writes into w by itself or returns []byte with info.
		// everything comes in in JSON body.
		info, err := action(w, rdr)

		if err != nil {
			logfile.Printf("%s", err)
			// serves error
			w.WriteHeader(http.StatusBadRequest)
			w.Write(makeresponse(err.Error(), string(info), []byte{}))
		}

	}
}

func action(w io.Writer, toreadbytes io.ReadCloser) ([]byte, error) {

	var inputBody []byte // data read from request body
	inputBody, err := ioutil.ReadAll(toreadbytes)
	if err != nil {
		logfile.Printf("%s", err)

		return []byte{}, err //this error goes to clients
	}

	inputStru, err := validate_input(inputBody)
	if err != nil {
		return []byte{}, err //this error goes to clients
	}

	// CreateDocxFromStruct creates docx files and writes them into w.
	info, err := CreateDocxFromStruct(w, inputStru, currentConfig.pathToTemplates)
	if err != nil {
		// may be an information message from CreateDocxFromStruct if showFileds parameter
		logfile.Printf("error: failed to create %s, %s\n", inputStru.DocxTemplateName, err)
		return info, err
	}
	logfile.Printf("successfully created %s\n", inputStru.DocxTemplateName)
	return []byte{}, nil
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
	sb.WriteString(`Help message: a service expects a JSON input in the following form
	{
		"DocxTemplateName": "youtemplatenamehere.docx",         //nonempty - a document template name
		"ShowFields": false, 									//help if you need a list of available MergeFields in document template
		"Header": {
			"Номер": "ЮХ000000084",								// these are MergeFields in the document template
			"Дата": "19.08.2021 11:31:20",                      // they are shown as example
			"КлиентПолноеНаименование": "?",
			"МестоСоставления": "?",
			"НомерДоговора": "?",
			"ДатаДоговора": "?",
			"ИтогоСумма": "645,26"
		},
		"Table1": [
			{
				// text 1, 2 ... in columns in you template document are just text, not MergeFields, in any row of any table in the document template
				"1": "1",
				"2": "Ремкомплект MM для редуктора СО2 Premium",
				"3": "шт",
				"4": "2",
				"5": "322,63",
				"6": "645,26"
			}
            // you may specify more rows as an input
		]
	}
	You must specify DocxTemplateName.
	You may use the following DocxTemplateName values:
	`)
	docs, err := alltemplates()
	if err != nil {
		sb.WriteString("The list of available document templates was not generated due to internal error.")
	} else {
		for k, _ := range docs {
			sb.WriteString(docs[k].Name())
		}
	}
	return sb.String()
}

func validate_input(databytes []byte) (*input, error) {
	inputStru, err := convertIntoinput(databytes) //converts json to struct
	if err != nil {
		return nil, errors.New(helpText())
	}

	if inputStru.DocxTemplateName == "" {

		return nil, errors.New(helpText())
	}
	return &inputStru, nil
}

// CreateDocxFromStruct creates docx document through gooxml, fills MergeFields with data from input JSON
// and adds rows into tables in docx document. Tables are searched by the row content: text "1" "2" in first
// two cells.
// Rows of Tables are filled from databytes which are json utf8 encoded struct "input".
func CreateDocxFromStruct(w io.Writer, inputStru *input, pathToTemplates string) ([]byte, error) {

	doc, err := getDocumentPtrFromTemplate(inputStru.DocxTemplateName, pathToTemplates)
	if err != nil {
		logfile.Printf("error: failed to get Document from template %s, %s\n", inputStru.DocxTemplateName, err)
		return nil, err
	}

	const specialMergeFields = "0123456789101112131415161718192021222324252627282930"
	helpmessage := make([]string, 0, 20)
	// first, merge fields from template documet will be filled with predefined value "__________"
	for _, v := range doc.MergeFields() {

		if inputStru.ShowFields {
			helpmessage = append(helpmessage, v)
		}
		if strings.Contains(specialMergeFields, v) {
			// if user made MergeFields in template with names like 1,2,3 ...
			// replace this merge fields with just text 1,2,3...
			inputStru.Header[v] = v
		}

		if _, has := inputStru.Header[v]; !has {
			// if user doesn't supply a value for the mergefield we will make it more visible
			// with default value for it
			inputStru.Header[v] = constUnderscore
		}
	}
	if inputStru.ShowFields { // client requested help for available MergeFields
		sort.Strings(helpmessage)
		return []byte(strings.Join(helpmessage, "; ")), errors.New("list of available MergeFields in document")
	}

	doc.MailMerge(inputStru.Header) // inserts values into the document by MailMerge

	if inputStru.Table1 != nil && len(inputStru.Table1) != 0 {
		// WORKING WITH TABLE
		// searches the table by row content "1 2 3 4 5"
		var tabfound bool
		var tabindex int
		var totalcells int // how many cells are in fact in the table in the document template

		tabfound, tabindex, totalcells = findOurTable(doc) // searches the table

		if !tabfound {
			err := errors.New("template file doesn't have a Table object with a row with 1,2,3 values in its cells.")
			logfile.Printf("%s", err)
			return []byte{}, err
		}
		tab := doc.Tables()[tabindex]

		//why? tab.Properties().SetStyle("TableGridZa") // I use style "TableGridZa"

		addRowWithCellsAndFillTexts(&tab, totalcells, inputStru.Table1) // adding rows
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
func findOurTable(doc *document.Document) (bool, int, int) {

	var tabfound bool
	var tabindex int
	var totalcells int // actual number of cells in the row

	tables := doc.Tables()
	for i, tab := range tables {

		rows := tab.Rows()
		for _, row := range rows {
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
				break
			}
		}

	}
	return tabfound, tabindex, totalcells

}

// we add a new method to an external package type Table
type TableWithDelete struct {
	*document.Table
}

// our additional method RemoveRow
func (t *TableWithDelete) RemoveRow(r document.Row) {
	for i, rc := range t.X().EG_ContentRowContent {

		if len(rc.Tr) > 0 && r.X() == rc.Tr[0] {

			copy(t.X().EG_ContentRowContent[i:], t.X().EG_ContentRowContent[i+1:])
			t.X().EG_ContentRowContent = t.X().EG_ContentRowContent[:len(t.X().EG_ContentRowContent)-1]
		}
	}
}

// addRowWithCellsAndFillTexts adds rows into tab from slice or rows (maps)
func addRowWithCellsAndFillTexts(tab *document.Table, totalcells int, sliceofmaps []map[int]string) {
	for _, datamap := range sliceofmaps {

		nrow := tab.AddRow()
		for nc := 1; nc <= totalcells; nc++ {

			ncell := nrow.AddCell()
			npar := ncell.AddParagraph()

			nrun := npar.AddRun()
			nrun.AddText(datamap[nc]) // nc is a column number passed in incoming json
		}

	}
}

func runHTTP(bindAddressPort string) error {
	var hand http.HandlerFunc = handlerhttp
	//starts serving
	err := http.ListenAndServe(bindAddressPort, hand)
	return err

}

// works with Windows Service Control Manager

// Tservice represents my service and has a method Execute
type Tservice struct {
	currentConfig serviceConfig
}

// Execute responds to SCM
func (s *Tservice) Execute(args []string, changerequest <-chan svc.ChangeRequest, updatestatus chan<- svc.Status) (ssec bool, errno uint32) {
	updatestatus <- svc.Status{State: svc.StartPending}

	//go runHTTP(s.currentConfig.bindAddressPort)

	supports := svc.AcceptStop | svc.AcceptShutdown

	updatestatus <- svc.Status{State: svc.Running, Accepts: supports}
	// select has no default and wait indefinitly
	select {
	case c := <-changerequest:
		switch c.Cmd {
		case svc.Stop, svc.Shutdown:
			goto stoped
		case svc.Interrogate:

		}
	}
stoped:
	return false, 0
}

// When Word saves a document, it removes all unused styles.  This means to
// copy the styles from an existing document, you must first create a
// document that contains text in each style of interest.
// for _, s := range doc.Styles.Styles() {
// 	fmt.Println("style", s.Name(), "has ID of", s.StyleID(), "type is", s.Type())
// }
