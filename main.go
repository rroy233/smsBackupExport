package main

import (
	"encoding/json"
	"fmt"
	"github.com/xuri/excelize/v2"
	"io/ioutil"
	"log"
	"net/http"
	"strconv"
	"time"
)

type httpRes struct {
	ErrorCode int    `json:"error_code,omitempty"`
	ErrorMsg  string `json:"error_msg,omitempty"`
	RequestId int64  `json:"request_id,omitempty"`
	TotalCnt  int    `json:"total_cnt"`
	List      []struct {
		ThreadID        string   `json:"thread_id"`
		Lctime          string   `json:"lctime"`
		LctimeMs        string   `json:"lctime_ms"`
		Unread          string   `json:"unread"`
		Box             string   `json:"box"`
		Person          string   `json:"person"`
		PersonFormatted string   `json:"person_formatted"`
		Name            string   `json:"name"`
		Card            string   `json:"card"`
		Content         string   `json:"content"`
		Type            string   `json:"type"`
		Md5             string   `json:"md5"`
		Imei            []string `json:"imei"`
		Ctime           string   `json:"_ctime"`
		Mtime           string   `json:"_mtime"`
		Isdelete        string   `json:"_isdelete"`
		Gmid            string   `json:"gmid"`
	} `json:"list"`
	RequestID int64 `json:"request_id"`
}

var styleMid int
var box = 0
var boxesUrl = []string{"receive", "send"}
var boxesName = []string{"接收", "发送"}

func main() {

	defer func() {
		log.Println("程序将在5s后结束运行...")
		time.Sleep(5 * time.Second)
	}()

	//读入cookie
	cookie, err := ioutil.ReadFile("./cookie.txt")
	if err != nil {
		log.Println(err)
		log.Println("读取cookie失败，请在可执行文件同目录下创建cookie.txt并填写")
		return
	}

	a := 0
	count := 0
	devices := make(map[string]int, 0)
	f := excelize.NewFile()

	//初始化样式
	styleMid, err = f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{
			Horizontal: "center",
		},
	})
	//填入导出信息
	f.SetCellValue("Sheet1", "A1", "导出时间")
	f.SetCellValue("Sheet1", "A2", time.Now().Format("2006-01-02 15:04:05"))
	if err != nil {
		log.Println(err)
		return
	}
	for {
		if box > 1 {
			//结束
			break
		}
		req, err := http.NewRequest("GET", fmt.Sprintf("https://duanxin.baidu.com/rest/2.0/pim/message?method=list&box=%s&app_id=20&imei=&card=&limit=%d-%d&t=", boxesUrl[box], a, a+99), nil)
		if err != nil {
			log.Println(err)
			return
		}
		req.Header.Set("Accept", "application/json, text/javascript, */*; q=0.01")
		req.Header.Set("Accept-Language", "zh-CN,zh-TW;q=0.9,zh;q=0.8")
		req.Header.Set("Connection", "keep-alive")
		req.Header.Set("Cookie", string(cookie))
		req.Header.Set("Dnt", "1")
		req.Header.Set("Referer", "https://duanxin.baidu.com/")
		req.Header.Set("Sec-Fetch-Dest", "empty")
		req.Header.Set("Sec-Fetch-Mode", "cors")
		req.Header.Set("Sec-Fetch-Site", "same-origin")
		req.Header.Set("User-Agent", "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/103.0.0.0 Safari/537.36")
		req.Header.Set("X-Requested-With", "XMLHttpRequest")
		req.Header.Set("Sec-Ch-Ua", "\".Not/A)Brand\";v=\"99\", \"Google Chrome\";v=\"103\", \"Chromium\";v=\"103\"")
		req.Header.Set("Sec-Ch-Ua-Mobile", "?0")
		req.Header.Set("Sec-Ch-Ua-Platform", "\"macOS\"")

		resp, err := http.DefaultClient.Do(req)
		if err != nil {
			log.Println(err)
			return
		}
		defer resp.Body.Close()

		data, err := ioutil.ReadAll(resp.Body)
		if err != nil {
			log.Println(err)
			return
		}

		res := new(httpRes)
		err = json.Unmarshal(data, res)
		if err != nil {
			log.Println(err)
			return
		}

		//判断是否错误
		if res.ErrorCode != 0 {
			log.Println("请求API发生错误:" + res.ErrorMsg)
			return
		}

		//判断是否切换
		if len(res.List) == 0 {
			a = 0
			log.Printf("已完成【%s】的导出\n", boxesName[box])
			time.Sleep(1 * time.Second)
			box++
			continue
		}

		//数据处理
		for _, msg := range res.List {
			//判断设备是否已记录
			deviceName := ""
			if len(msg.Imei[0]) > 33 {
				deviceName = msg.Imei[0][:len(msg.Imei[0])-33]
			} else {
				deviceName = msg.Imei[0]
			}
			if devices[deviceName] == 0 {
				devices[deviceName] = 2
				addHeader(f, deviceName)
			}
			//输出到表格
			name := ""
			if msg.PersonFormatted != msg.Name {
				name = msg.Name
			}
			data := map[string]interface{}{
				"A": devices[deviceName] - 1,
				"B": boxesName[box],
				"C": deviceName,
				"D": "-",
				"E": "-",
				"F": "-",
				"G": "-",
				"H": msg.Content,
				"I": formatTs(msg.Lctime),
			}
			if box == 0 {
				//接受
				data["D"] = name
				data["E"] = msg.PersonFormatted
			} else {
				//发送
				data["F"] = name
				data["G"] = msg.PersonFormatted
			}
			for k, v := range data {
				err = f.SetCellValue(deviceName, fmt.Sprintf("%s%d", k, devices[deviceName]), v)
				if err != nil {
					log.Println(err)
					return
				}
			}
			devices[deviceName]++

			log.Printf("[%d][%s][%s]%s\n", count, boxesUrl[box], deviceName, msg.Content)
			count++
		}
		//time.Sleep(500 * time.Millisecond)
		a += 100
	}

	if err := f.SaveAs("out.xlsx"); err != nil {
		log.Println(err)
		return
	}
	log.Println("已保存为out.xlsx")
}

func addHeader(file *excelize.File, sheetName string) {
	header := map[string]string{
		"A1": "ID",
		"B1": "收/发",
		"C1": "设备",
		"D1": "发件人姓名",
		"E1": "发件人号码",
		"F1": "收件人姓名",
		"G1": "收件人号码",
		"H1": "短信内容",
		"I1": "时间",
	}
	file.NewSheet(sheetName)
	for k, v := range header {
		err := file.SetCellValue(sheetName, k, v)
		if err != nil {
			log.Println(err)
			return
		}
		file.SetColWidth(sheetName, "A", "B", 6)
		file.SetColWidth(sheetName, "C", "G", 20)
		file.SetColWidth(sheetName, "H", "H", 140)
		file.SetColWidth(sheetName, "I", "I", 17)
		file.SetColStyle(sheetName, "A:I", styleMid)
	}
	return
}

func formatTs(ts string) string {
	t, err := strconv.ParseInt(ts, 10, 64)
	if err != nil {
		return ts
	}
	return time.Unix(t/1000, 0).Format("2006-01-02 15:04:05")
}
