package main

import (
	"bytes"
	"context"
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"fmt"
	"log"
	"net/http"
	"time"

	"github.com/aws/aws-sdk-go-v2/aws"
	v4 "github.com/aws/aws-sdk-go-v2/aws/signer/v4"
	"github.com/xuri/excelize/v2"
)

type Payload struct {
	Payload map[string]interface{} `json:"payload"`
}

type User struct {
	Nombre      string `json:"nombre"`
	Code        int    `json:"code"`
	Username    string `json:"username"`
	IdeCanal    string `json:"ideCanal"`
	IdeVendedor string `json:"ideVendedor"`
	UsuEmision  string `json:"usuEmision"`
}

type Trace struct {
	Id     string `json:"id"`
	ApiKey string `json:"apiKey"`
}

type ErrorApi struct {
	ErrorApi map[string]interface{} `json:"error"`
}

func main() {
	accessKey := "ACCESS_KEY"
	secretKey := "SECRET_KEY"
	service := "SERVICE"
	region := "REGION"
	url := "YOUR_URL_API"

	creds := aws.Credentials{
		AccessKeyID: accessKey, SecretAccessKey: secretKey, Source: "manual",
	}
	filePath := "./your_path_file.xlsx"
	sheetName := "sheetname"

	f, err := excelize.OpenFile(filePath)
	if err != nil {
		log.Fatalf("Error al abrir el archivo Excel: %v", err)
	}
	defer f.Close()

	rows, err := f.GetRows(sheetName)
	if err != nil {
		log.Fatalf("Error al obtener las filas de la hoja: %v", err)
	}

	signer := v4.NewSigner()

	for i, row := range rows {
		if len(row) >= 9 {
			data := row[8] // Columna H (índice 7)

			var payload Payload
			err := json.Unmarshal([]byte(data), &payload)
			if err != nil {
				log.Printf("Error al parsear JSON en fila %d: %v", i+1, err)
				continue
			}

			payload.Payload["producto"] = "PRODUCTO"
			payload.Payload["user"] = User{
				Nombre:      "XXXX-XXXX",
				Code:        123456,
				Username:    "XXXX-XXXX",
				IdeCanal:    "XXXX-XXXX",
				IdeVendedor: "XXXX-XXXX",
				UsuEmision:  "XXXX-XXXX",
			}

			payload.Payload["trace"] = Trace{
				Id:     "XXXX-XXXX",
				ApiKey: "XXXX-XXXX",
			}

			formattedJSON, err := json.MarshalIndent(payload, "", "  ")
			if err != nil {
				log.Printf("Error al formatear JSON en fila %d: %v", i+1, err)
				continue
			}
			fmt.Printf("JSON en fila %d:\n%s\n", i+1, string(formattedJSON))

			jsonData, err := json.Marshal(payload)
			if err != nil {
				log.Printf("Error en el marshal JSON para la fila %d: %v", i+1, err)
				continue
			}

			req, err := http.NewRequest("PATCH", url, bytes.NewBuffer(jsonData))
			if err != nil {
				log.Printf("Error al crear la peticion http en la fila %d: %v", i+1, err)
				continue
			}

			req.Header.Set("Content-Type", "application/json")
			req.Header.Set("x-api-key", "XXXX-XXXX")
			req.Header.Set("traceId", "XXXX-XXXX")

			hash := sha256.New()
			hash.Write(jsonData)
			payloadHash := hex.EncodeToString(hash.Sum(nil))

			err = signer.SignHTTP(context.TODO(), creds, req, payloadHash, service, region, time.Now())
			if err != nil {
				log.Printf("Error al firmar la peticion en la fila %d: %v", i+1, err)
				continue
			}

			client := &http.Client{}
			resp, err := client.Do(req)
			if err != nil {
				log.Printf("Error al realizar la peticion en la fila %d: %v", i+1, err)
				continue
			}
			defer resp.Body.Close()

			buf := new(bytes.Buffer)
			buf.ReadFrom(resp.Body)
			responseBody := buf.String()

			var responseApi ErrorApi
			errorApi := json.Unmarshal([]byte(responseBody), &responseApi)
			if errorApi != nil {
				log.Printf("Error al parsear JSON en fila %d: %v", i+1, errorApi)
				continue
			}

			responseFormatted, err := json.MarshalIndent(responseApi, "", "  ")
			if err != nil {
				log.Printf("Error al formatear JSON en fila %d: %v", i+1, err)
				continue
			}

			fmt.Printf("Respuesta en fila %d:\n%s\n", i+1, responseFormatted)

			cell := fmt.Sprintf("M%d", i+1)
			f.SetCellValue(sheetName, cell, responseFormatted)
		}
	}

	if err := f.SaveAs("./your_path_file.xlsx"); err != nil {
		log.Fatalf("Error al guardar el archivo Excel: %v", err)
	}

	fmt.Println("Proceso completado!!!")
}
