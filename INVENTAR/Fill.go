// to debug run
// curl.exe -T .\ex1.txt --url http://127.0.0.1:1313/docxcreator --verbose --output .\out2.docx
package main

import (
	"encoding/json"
	"errors"
	"flag"
	"fmt"
	"io"
	"io/ioutil"
	"log"
	"net/http"
	"os"
	"path/filepath"
	"sort"
	"strings"

	"baliance.com/gooxml/document"
)

// zaHeaderAndTable we expecting as json
type zaHeaderAndTable struct {
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
}

// error response struct, returned by service as json
type responseStruct struct {
	Error   string
	Data    []byte
	StrData string
}

var currentConfig serviceConfig
var logfile *log.Logger

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
	pathToTemplates := flag.String("PathToTemplates", "", "path to templates files. Template name expected in incoming json.")
	jsonFileName := flag.String("jsonFileName", "", `file with jsoned data, utf-8. Not-a-service mode. Example of json: {"Header":{"key1":"val1"},"Table1":[{"#numberOfColumn":"valueinColumnnOfTable1"}]}`)

	bindAddressPort := flag.String("bindAddressPort", "127.0.0.1:8080", "bind service to address and port")

	fullPathLogFile := flag.String("logfile", "", "path and name to service log file")
	flag.Parse()

	if *fullPathLogFile == "" {
		flag.Usage()
		log.Fatalf("%s\n", "не передан параметр logfile - полный путь к лог файлу")
	}

	flog, err := os.OpenFile(*fullPathLogFile, os.O_RDWR|os.O_CREATE, 0)
	if err != nil {
		log.Fatalf("%s\n%s\n", "Error: can't create log file: "+(*fullPathLogFile), err)
	}

	logfile = log.New(flog, "", log.Ldate|log.Ltime)

	// log is ready

	if *pathToTemplates == "" {
		flag.Usage()
		log.Fatalf("%s\n", "Error: PathToTemplates - path to a folder with docx templates needed.")
	}

	if *jsonFileName == "" { //starts as a service, uses http

		// global var
		currentConfig = serviceConfig{
			pathToTemplates: *pathToTemplates,
			fullPathLogFile: *fullPathLogFile,
		}

		var hand http.HandlerFunc = handlerhttp
		//starts serving
		err = http.ListenAndServe(*bindAddressPort, hand)

		fmt.Printf("%s\n%s", "Http server Exited:", err)

	} else { //command line mode, expecting file
		f, err := os.Open(*jsonFileName)
		if err != nil {
			log.Fatalf("%s\n%s\n", err, "Error: can't open file with json utf-8.")
		}
		databytes, err := ioutil.ReadAll(f)

		новфайл := "debugfile.docx" //fixed output docx file name
		newfilefullpath := filepath.Join(".\\", новфайл)
		wdebug, err := os.OpenFile(newfilefullpath, os.O_CREATE|os.O_WRONLY, 0) //for debug
		if err != nil {
			logfile.Printf("%s\n%s %s", err, "Error: can't output to file", newfilefullpath)
			os.Exit(1)
		}

		info, err := CreateDocxFromStruct(wdebug, databytes, *pathToTemplates)
		if err != nil {
			fmt.Printf("%s\n%s", err, string(info))
		}
	}

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

		// the action on /docxcreator url writes into w by itself or returns []byte with info
		info, err := action(w, rdr)
		if err != nil {
			logfile.Printf("%s", err)
			// serves error
			w.WriteHeader(http.StatusBadRequest)
			w.Write(makeresponse(fmt.Sprintf("%s", err), string(info), []byte{}))
		}

	}
}

func action(w io.Writer, toreadbytes io.ReadCloser) ([]byte, error) {

	var databytes []byte // data read from request body
	databytes, err := ioutil.ReadAll(toreadbytes)
	if err != nil {
		logfile.Printf("%s", err)
		return []byte{}, err
	}
	// CreateDocxFromStruct creates docx files (writes to w)
	info, err := CreateDocxFromStruct(w, databytes, currentConfig.pathToTemplates)
	if err != nil {
		// may be an information message from CreateDocxFromStruct if showFileds parameter
		logfile.Printf("%s\n%s\n", "Error: failed create docx file.", err)
		return info, err
	}
	return []byte{}, nil
}

// CreateDocxFromStruct creates doxc document through gooxml and fills mergefields
// and adds rows into a table. Table searched by the row content: "1 2 3 4 5"
// Rows filled from databytes which are json utf8 encoded struct zaHeaderAndTable.
func CreateDocxFromStruct(w io.Writer, databytes []byte, pathToTemplates string) ([]byte, error) {

	datastr, err := getzaHeaderAndTable(databytes) //converts json to struct
	if err != nil {
		logfile.Printf("%s\n%s\n", "Error: your sent json data parsing fails.", err)
		return []byte{}, err
	}

	if datastr.DocxTemplateName == "" {
		errstr := "Your json data must contain DocxTemplateName field, the name of a template."
		logfile.Printf("%s\n", errstr)
		return []byte{}, errors.New(errstr)
	}

	//opens template
	doc, err := document.Open(filepath.Join(pathToTemplates, datastr.DocxTemplateName))
	if err != nil {
		flag.Usage()
		logfile.Printf("%s\n%s", err, "Error: can't read docx template file.")
		return []byte{}, err
	}
	// When Word saves a document, it removes all unused styles.  This means to
	// copy the styles from an existing document, you must first create a
	// document that contains text in each style of interest.
	//Used style is "TableGridZa"
	// for _, s := range doc.Styles.Styles() {
	// 	fmt.Println("style", s.Name(), "has ID of", s.StyleID(), "type is", s.Type())
	// }
	helpmessage := make([]string, 0, 20)
	// merge fields from template documet will be filled with predefined value "________"
	for _, v := range doc.MergeFields() {
		if datastr.ShowFields {
			helpmessage = append(helpmessage, v)
		}
		_, is := datastr.Header[v]
		if !is {
			datastr.Header[v] = "____________"
		}
	}
	if datastr.ShowFields {
		sort.Strings(helpmessage)
		return []byte(strings.Join(helpmessage, "; ")), errors.New("just info")
	}

	doc.MailMerge(datastr.Header) // inserts Header values into document by MailMerge

	// WORKING WITH TABLE
	// searches the table by row content "1 2 3 4 5"
	var tabfound bool
	var tabindex int
	var totalcells int // how many cells are in fact in the table in the document template

	tabfound, tabindex, totalcells = findOurTable(doc) // searches the table

	if tabfound {
		tab := doc.Tables()[tabindex]
		tab.Properties().SetStyle("TableGridZa") // I use style "TableGridZa"

		addRowWithCellsAndFillTexts(&tab, totalcells, datastr.Table1) // adding rows
	}

	// saves new dowcument into io.Writer
	err = doc.Save(w)
	if err != nil {
		logfile.Printf("%s", "Error: can't write into io.writer.")
		return []byte{}, err
	}
	return []byte{}, nil
}

// getzaHeaderAndTable unmarshals a json into struct
func getzaHeaderAndTable(bstr []byte) (zaHeaderAndTable, error) {
	var v zaHeaderAndTable
	err := json.Unmarshal(bstr, &v)
	if err != nil {
		log.Printf("%s\n%s\n", "Error: can't use your json.", err)
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
