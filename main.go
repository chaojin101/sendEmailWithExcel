package main

import (
	"fmt"
	"log"
	"strconv"
	"strings"

	"github.com/chaojin101/comUtils"
	"github.com/chaojin101/go-email"
	"github.com/xuri/excelize/v2"
)

func main() {
	f, err := excelize.OpenFile("config.xlsx")
	if err != nil {
		fmt.Println(err)
		comUtils.Input()
		return
	}
	defer func() {
		// Close the spreadsheet.
		if err := f.Close(); err != nil {
			fmt.Println(err)
			comUtils.Input()
		}
	}()
	// Get all the rows in the Sheet1.
	rows, err := f.GetRows("Sheet1")
	if err != nil {
		fmt.Println(err)
		comUtils.Input()
		return
	}

	fmt.Print("Start line: ")
	startLine, _ := strconv.Atoi(comUtils.Input())
	fmt.Print("End line: ")
	endLine, _ := strconv.Atoi(comUtils.Input())
	if startLine < 1 || endLine > len(rows) {
		log.Println(startLine, endLine, "error")
		comUtils.Input()
		return
	}

	for i, row := range rows[startLine-1 : endLine] {

		s := email.Sender{
			Name:            row[0],
			Mail:            row[1],
			MailSmtpAddress: row[2],
			MailAuthcode:    row[3],
			Subject:         row[6],
			Text:            row[7],
			Attach:          strings.Split(row[8], " "),
		}

		r := email.Recipient{
			Name: row[4],
			Mail: row[5],
		}

		msg := "sent successfully."
		err := s.Send(r)
		if err != nil {
			msg = "error: " + err.Error()
		}
		log.Println(i+startLine, s.Name, s.Mail, r.Name, r.Mail, s.Subject, msg)

	}
	log.Println("Over.")
	comUtils.Input()
}
