package main

import (
	"flag"
	"fmt"
	"time"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
)

var isDebug = flag.Int("debug", 0, "is debug")

func main() {
	flag.Parse()
	fmt.Println(*isDebug)

	defer func() {
		fmt.Println("\n20秒后自动关闭...")
		time.Sleep(10 * time.Second)
	}()

	fmt.Println("1.读取文件")
	f, err := excelize.OpenFile("1.xlsx")
	if err != nil {
		fmt.Printf("ERROR:文件读取错误,请把文件重命名为1.xlsx,放入到软件同一目录下.%v", err)
		return
	}

	fee := NewFee(f, *isDebug != 0)
	if err := fee.Start(); err != nil {
		return
	}

	fmt.Println("6.保存文件")
	if err := f.SaveAs("1-ok.xlsx"); err != nil {
		fmt.Printf("ERROR:保存文件失败,%v", err)
		return
	}

	fmt.Println("SUCCESS")
	fmt.Println("用excel打开后-记得点击-公式-重算工作簿按钮")
}
