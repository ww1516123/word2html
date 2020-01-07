package main

import (
	"bufio"
	"encoding/base64"
	"flag"
	"fmt"
	"io"
	"os"
	"strconv"
	"strings"

	"github.com/axgle/mahonia"
)

var filePath string
var encoding string

type metaInfo struct {
	from        string
	subject     string
	version     float64
	flag        string
	contentType string
	boundary    string
}

// 内容Meta信息
type boundaryMeta struct {
	contentType             string
	contentCharset          string
	contentTransferEncoding string
	contentDisposition      string
	contentID               string
}

type imageInfo struct {
	imageMeta    boundaryMeta
	imageContent string
}

type docInfo struct {
	body        string
	images      map[string]imageInfo
	readMeta    bool
	bodyContent boundaryMeta
	fileMeta    metaInfo
}

func main() {
	flag.StringVar(&filePath, "f", "", "filePath")
	flag.StringVar(&encoding, "e", "GBK", "encoding")
	flag.Parse() //暂停获取参数
	if filePath == "" {
		fmt.Printf("Error: %s\n", "please input filePath")
		return
	}
	readFile(filePath)
}

/**
 * 文件读取
 **/
func readFile(path string) {
	fi, err := os.Open(path)
	if err != nil {
		fmt.Printf("Error: %s\n", err)
		return
	}
	defer fi.Close()
	decoder := mahonia.NewDecoder("gb18030")
	br := bufio.NewReader(fi)
	metaInfo := metaInfo{}
	boundaryMeta := boundaryMeta{}
	docInfo := docInfo{}
	docInfo.images = make(map[string]imageInfo, 1)
	i := 0
	for {
		line, _, c := br.ReadLine()
		if c == io.EOF {
			break
		}
		lineStr := string(line)
		decoderStr := decoder.ConvertString(lineStr)
		if i < 5 {
			getmetaInfo(&metaInfo, decoderStr)
			i++
		} else {
			readBoundary(metaInfo, &boundaryMeta, &docInfo, decoderStr)
		}
	}
	path = strings.ReplaceAll(path, ".docx", ".html")
	path = strings.ReplaceAll(path, ".doc", ".html")
	write2file(&docInfo, path)
}

/**
 * 判断文件是否存在  存在返回 true 不存在返回false
 */
func checkFileIsExist(filename string) bool {
	var exist = true
	if _, err := os.Stat(filename); os.IsNotExist(err) {
		exist = false
	}
	return exist
}

/**
 * 判断文件是否存在  存在返回 true 不存在返回false
 */
func write2file(docInfo *docInfo, filename string) {
	if docInfo.bodyContent.contentCharset == "gb2312" {
		docInfo.bodyContent.contentCharset = "GBK"

	}
	if encoding == ""{
		encoding = docInfo.bodyContent.contentCharset
	}

	decoder := mahonia.NewDecoder(encoding)
	decodeBytes, _ := base64.StdEncoding.DecodeString(docInfo.body)
	var f *os.File
	//如果文件存在
	if checkFileIsExist(filename) {
		//打开文件
		f, _ = os.OpenFile(filename, os.O_WRONLY, 0777)
	} else {
		//创建文件
		f, _ = os.Create(filename)
	}
	//创建新的 Writer 对象
	w := bufio.NewWriter(f)
	defer f.Close()
	html := decoder.ConvertString(string(decodeBytes))
	//编码集
	html = strings.Replace(html, "<html>", "<html><meta charset=\"UTF-8\">", 1)
	// fmt.Printf(" the %v",docInfo)
	for key, value := range docInfo.images {

		html = strings.Replace(html, "src=\"cid:"+key+"\"", "src=\"data:"+value.imageMeta.contentType+";base64,"+value.imageContent+"\"", 1)
	}
	w.Write([]byte(html))
	w.Flush()
}

/**
 * 读取boundary 信息
 **/
func readBoundary(metaInfo metaInfo, boundaryMetaInfo *boundaryMeta, docInfo *docInfo, decoderStr string) {
	// fmt.Printf("the readMeta %t", docInfo.readMeta)
	if docInfo.readMeta {
		decoderStr = strings.ReplaceAll(decoderStr, " ", "")
		if len(decoderStr) == 0 {
			docInfo.readMeta = false
			if boundaryMetaInfo.contentType == "text/html" {
				bodyContent := boundaryMeta{}
				bodyContent.contentCharset = boundaryMetaInfo.contentCharset
				bodyContent.contentTransferEncoding = boundaryMetaInfo.contentTransferEncoding
				bodyContent.contentType = boundaryMetaInfo.contentType
				docInfo.bodyContent = bodyContent
			}
			return
		}
		getBodyMeta(boundaryMetaInfo, decoderStr)
	} else {
		if strings.Index(decoderStr, metaInfo.boundary) >= 0 {
			docInfo.readMeta = true
			return
		}
		if boundaryMetaInfo.contentType == "text/html" {
			docInfo.body += decoderStr
		}
		if boundaryMetaInfo.contentType == "image/png" || boundaryMetaInfo.contentType == "image/gif" {
			imageInfo := imageInfo{}
			imageStr := ""
			if _, have := docInfo.images[boundaryMetaInfo.contentID]; have {
				imageInfo = docInfo.images[boundaryMetaInfo.contentID]
				imageStr = imageInfo.imageContent
			} else {
				imageContent := boundaryMeta{}
				imageContent.contentCharset = boundaryMetaInfo.contentCharset
				imageContent.contentTransferEncoding = boundaryMetaInfo.contentTransferEncoding
				imageContent.contentType = boundaryMetaInfo.contentType
				imageContent.contentID = boundaryMetaInfo.contentID
				imageContent.contentDisposition = boundaryMetaInfo.contentDisposition
				imageInfo.imageMeta = imageContent
			}
			imageStr += decoderStr
			imageInfo.imageContent = imageStr
			docInfo.images[boundaryMetaInfo.contentID] = imageInfo
		}
	}
}

/**
 * 读取content Meta信息
 */
func getBodyMeta(boundaryMeta *boundaryMeta, line string) {
	if strings.Index(line, "Content-Type") >= 0 {
		lines := strings.Split(strings.Split(line, ":")[1], ";")
		boundaryMeta.contentType = lines[0]
		if len(lines) > 1 {
			boundaryMeta.contentCharset = strings.ReplaceAll(strings.Split(lines[1], "=")[1], "\"", "")
		}
	} else if strings.Index(line, "Content-Transfer-Encoding") >= 0 {
		boundaryMeta.contentTransferEncoding = strings.Split(line, ":")[1]
	} else if strings.Index(line, "Content-Disposition") >= 0 {
		boundaryMeta.contentDisposition = strings.Split(line, ":")[1]
	} else if strings.Index(line, "Content-ID") >= 0 {
		boundaryMeta.contentID = strings.ReplaceAll(strings.ReplaceAll(strings.Split(line, ":")[1], "<", ""), ">", "")
	}
}

/**
 * 读取源文件的meta信息
 */
func getmetaInfo(metaInfo *metaInfo, line string) {
	i := strings.Index(line, "From")
	if i >= 0 {
		metaInfo.from = strings.Split(line, ":")[1]
	} else if strings.Index(line, "Subject") >= 0 {
		metaInfo.subject = strings.Split(line, ":")[1]
	} else if strings.Index(line, "MIME-Version") >= 0 {
		thisVersion, _ := strconv.ParseFloat(strings.Split(line, ":")[1], 64)
		metaInfo.version = thisVersion
	} else if strings.Index(line, "X-51JOB-FLAG") >= 0 {
		metaInfo.flag = strings.Split(line, ":")[1]
	} else if strings.Index(line, "Content-Type") >= 0 {
		metaInfo.contentType = strings.Split(line, ":")[1]
		metaInfo.boundary = strings.ReplaceAll(strings.ReplaceAll(strings.Split(metaInfo.contentType, ";")[1], "boundary=", ""), "\"", "")
	}
}
