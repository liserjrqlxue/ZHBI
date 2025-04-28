package main

import (
	"embed"
	"fmt"
	"html/template"
	"io"
	"io/fs"
	"log"
	"log/slog"
	"math/rand"
	"strings"
	"sync"
	"time"

	// "log/slog"
	"net/http"
	"os"
	"path/filepath"

	"github.com/liserjrqlxue/goUtil/sge"
)

//go:embed templates/*.html
var templateFiles embed.FS

var (
	tasks   = make(map[string]*Task)
	taskMux sync.Mutex
)

type Task struct {
	ID        string
	Filename  string
	Param     string
	Workdir   string
	Basename  string
	Status    string
	Result    string // result content
	Completed bool
}

func uploadPage(w http.ResponseWriter, r *http.Request) {
	tmpl, err := ParseFileOrFS("templates/upload.html", templateFiles)
	if err != nil {
		log.Println("Error parsing the templates:", err)
		http.Error(w, err.Error(), http.StatusInternalServerError)
		return
	}
	tmpl.Execute(w, nil)
}

func startHandler(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Redirect(w, r, "/", http.StatusSeeOther)
		return
	}

	err := r.ParseMultipartForm(32 << 20)
	if err != nil {
		log.Println("Error parsing the form:", err)
		http.Error(w, err.Error(), http.StatusInternalServerError)
		return
	}

	file, handler, err := r.FormFile("file")
	if err != nil {
		http.Error(w, "file error", http.StatusBadRequest)
		return
	}
	defer file.Close()
	param := r.FormValue("param")

	taskID := fmt.Sprintf("%d", rand.Int63())
	date := time.Now().Format("20060102")
	baseName := strings.TrimSuffix(handler.Filename, ".xlsx")
	workdir := fmt.Sprintf("public/%s/%s", date, baseName)
	err = os.MkdirAll(workdir, 0755)
	input := fmt.Sprintf("%s/%s", workdir, handler.Filename)
	if err != nil {
		log.Println("Error mkdir:", workdir, err)
		http.Error(w, err.Error(), http.StatusInternalServerError)
		return
	}
	f, _ := os.Create(input)
	defer f.Close()
	io.Copy(f, file)

	task := &Task{
		ID:       taskID,
		Filename: handler.Filename,
		Param:    param,
		Status:   "Running",
		Workdir:  workdir,
		Basename: baseName,
	}
	slog.Info("Task", "task", task)

	taskMux.Lock()
	tasks[taskID] = task
	taskMux.Unlock()

	go processTask(task)

	http.Redirect(w, r, "/result?taskID="+taskID, http.StatusSeeOther)
}

func processTask(task *Task) {
	var args = []string{
		"-i",
		filepath.Join(task.Workdir, task.Filename),
	}
	if task.Param != "on" {
		args = append(args, "-s")
	}

	taskMux.Lock()
	err := sge.Run("../SynOrdEval/SynOrdEval.exe", args...)
	if err != nil {
		task.Status = "Failed"
		task.Result = fmt.Sprintf("分析失败: err=[%+v]", err)
	} else {
		task.Status = "Success"
		task.Result = fmt.Sprintf("处理完成: 输入文件=[%s], 参数=[%s]", task.Filename, task.Param)
	}
	task.Completed = true
	taskMux.Unlock()
}

func resultPage(w http.ResponseWriter, r *http.Request) {
	tmpl, err := ParseFileOrFS("templates/result.html", templateFiles)
	if err != nil {
		log.Println("Error parsing the templates:", err)
		http.Error(w, err.Error(), http.StatusInternalServerError)
		return
	}

	taskID := r.URL.Query().Get("taskID")
	task, ok := tasks[taskID]
	if !ok {
		task = &Task{ID: taskID}
	}
	tmpl.Execute(w, task)
}

func statusHandler(w http.ResponseWriter, r *http.Request) {
	taskID := r.URL.Query().Get("taskID")

	taskMux.Lock()
	task, ok := tasks[taskID]
	taskMux.Unlock()

	if !ok {
		http.Error(w, "task not found", http.StatusNotFound)
		return
	}

	if task.Completed {
		fmt.Fprintf(w, `{"done": true, "result": "%s"}`, task.Result)
	} else {
		fmt.Fprintf(w, `{"done": false, "status": %s}`, task.Status)
	}
}

func ParseFileOrFS(path string, fs fs.FS) (tmpl *template.Template, err error) {
	tmpl, err = template.ParseFiles(path)
	if err != nil {
		tmpl, err = template.ParseFS(fs, path)
	}
	return
}
