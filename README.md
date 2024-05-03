# MasadoraOrderCrawler
python爬虫爬取魔法集市全部代购订单

## 已失效注意
目前玛莎提供的订单接口已经失效，因此无法再通过此程序导出订单数据。本页面仅留做纪念

## 注意
- 仅用于对即将失效的订单记录进行备份

- 仅供个人使用，不对该程序的任何后续影响负责

- 作者代码水平很菜，有bug请在Issues里提，非常感谢！

## 使用方法：
- 下载[MasadoraOrdersCrawler.zip](https://github.com/GardenKitten/MasadoraOrderCrawler/releases/download/ver1.0/MasadoraOrdersCrawler.zip)压缩包并解压
- 打开玛莎订单页：https://www.masadora.jp/member/order
- 在该页面按Ctrl+Shift+I调出开发者工具，打开Network一栏
- 刷新玛莎订单页，在开发者工具中查看名为order的request
- 查看order的header中的Request Headers，找到里面的Cookie
- 打开压缩包中的config.yaml
- 将Cookie中的__afmk一项配置到yaml文件中并保存
- 运行MasadoraOrderCrawler.exe，等待程序运行完毕
- 如果弹出“Excel 文件生成成功！“的提示，则能在根目录中看到生成的订单excel文件
