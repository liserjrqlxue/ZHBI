package main

import (
	"flag"
	"log"
	"net"
	"net/http"
	"os"

	"github.com/liserjrqlxue/goUtil/simpleUtil"
)

// flag
var (
	port = flag.String("port", ":9091", "port")
)

func main() {
	flag.Parse()

	// mkdir work dir
	os.MkdirAll("public", 0755)

	// handle file serve
	http.Handle("/static/", http.StripPrefix("/static/", http.FileServer(http.Dir("static"))))
	http.Handle("/public/", http.StripPrefix("/public/", http.FileServer(http.Dir("public"))))

	http.HandleFunc("/", uploadPage)
	http.HandleFunc("/upload", uploadPage)
	http.HandleFunc("/start", startHandler)
	http.HandleFunc("/result", resultPage)
	http.HandleFunc("/status", statusHandler)

	log.Printf("start http://%v%v\n", GetOutboundIP(), *port)
	simpleUtil.CheckErr(http.ListenAndServe(*port, nil))
}

// Get preferred outbound ip of this machine
func GetOutboundIP() net.IP {
	conn, err := net.Dial("udp", "8.8.8.8:80")
	if err != nil {
		log.Fatal(err)
	}
	defer conn.Close()

	localAddr := conn.LocalAddr().(*net.UDPAddr)
	return localAddr.IP
}
