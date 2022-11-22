package main

import (
	"flag"
	"fmt"
	"github.com/gin-contrib/cors"
	"strconv"
	"time"

	"github.com/gin-gonic/gin"
	"github.com/mozillazg/go-pinyin"
)

var pinyinArgs pinyin.Args

func init() {
	pinyinArgs = pinyin.NewArgs()
	pinyinArgs.Style = pinyin.Tone
}

func createCorsConfig() cors.Config {
	return cors.Config{
		AllowOrigins:     []string{"*"},
		AllowMethods:     []string{"GET"},
		AllowHeaders:     []string{"Origin"},
		ExposeHeaders:    []string{"Content-Length"},
		AllowCredentials: true,
		AllowOriginFunc: func(origin string) bool {
			return origin == "*"
		},
		MaxAge: 12 * time.Hour,
	}
}

func runHttpServer(sPort string) {
	router := gin.Default()
	router.Use(cors.New(createCorsConfig()))
	router.GET("/pinyin", handleQueryPinyinList)
	router.GET("/pinyin1", handleQueryPinyin)
	_ = router.Run(sPort)
}

// Call like GET http://localhost:8080/pinyin?han=我来了
func handleQueryPinyinList(c *gin.Context) {
	han := c.DefaultQuery("han", "")

	p := pinyin.Pinyin(han, pinyinArgs)

	c.JSON(200, gin.H{"code": 0, "data": p})
}

// Call like GET http://localhost:8080/pinyin1?han=我
func handleQueryPinyin(c *gin.Context) {
	han := c.DefaultQuery("han", "")

	p := pinyin.Pinyin(han, pinyinArgs)

	s := ""
	if len(p) > 0 {
		s = p[0][0]
	}

	c.JSON(200, gin.H{"code": 0, "data": s})
}

func main() {
	fmt.Print("\n\nDEFAULT PORT: 8080, USING '-port portnum' TO START ANOTHER PORT.\n\n")

	port := flag.Int("port", 8080, "Port Number, default 8080")
	flag.Parse()
	sPort := ":" + strconv.Itoa(*port)

	runHttpServer(sPort)
}
