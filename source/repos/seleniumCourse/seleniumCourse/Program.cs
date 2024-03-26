using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Diagnostics;

class Program
{
    static void Main(string[] args)
    {
        // 设置 Driver 的路径
        string driverPath = @"C:\Users\Administrator\Desktop\chromedriver-win64";

        // 创建 Driver 实例
        ChromeOptions options = new ChromeOptions();
        options.AddArgument("--start-maximized"); // 最大化窗口
        IWebDriver driver = new ChromeDriver(driverPath, options);

        // 打开网页
        driver.Navigate().GoToUrl("https://wsoa.wicresoft.com/Home");

        // 找到用户名和密码的输入框并输入信息
        driver.FindElement(By.XPath("/html/body/div/div/div[2]/div/div/form/div/div/div[1]/div[1]/input")).SendKeys("ariess");
        driver.FindElement(By.XPath("/html/body/div/div/div[2]/div/div/form/div/div/div[1]/div[2]/input")).SendKeys("Szh008008!!");

        // 找到登录按钮并点击
        driver.FindElement(By.XPath("/html/body/div[1]/div/div[2]/div/div/form/div/div/div[2]/div[2]/button")).Click();

        Thread.Sleep(5000);
        driver.FindElement(By.XPath("/html/body/div[1]/div[3]/div/div[1]/ul/li[3]/a")).Click();
        // 等待30秒这个过程中选择课程进入播放器
        Thread.Sleep(30000);

        WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromMinutes(10)); // 显性等待
        Console.WriteLine(driver.WindowHandles[1]);
        var driver2 = driver;
        for (int i=0; i< driver.WindowHandles.Count();i++) 
        {
            
            Console.WriteLine(i);
            //打开播放器，保持浏览器只有3个页面，1.工作台，2.学习平台，3.播放器页面
            if(driver2.SwitchTo().Window(driver.WindowHandles[i]).Title != "工作台-微创软件办公平台" && driver2.SwitchTo().Window(driver.WindowHandles[i]).Title != "微创在线学习平台")
            { 
                driver.SwitchTo().Window(driver.WindowHandles[i]);
                Console.WriteLine("第" + i + "个网页    " + "网页名字：" + driver.Title);
                break;
            }
                        
        }

        driver.SwitchTo().Frame("aliPlayerFrame");
        //循环次数计数器
        int loop = 0;
        while (true)
        {
            // 执行JavaScript以设置视频播放速率
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            var a = js.ExecuteScript("document.querySelector('video').playbackRate = 11");
            while (true)
            {
                try
                {
                    var buttonElement = driver.FindElement(By.XPath("//*[@id='app']/div/div[2]/div[1]/div[6]/div/button"));
                    buttonElement.Click();
                    Thread.Sleep(TimeSpan.FromSeconds(7));
                    break;
                }
                catch (NoSuchElementException)
                {
                    Thread.Sleep(5000);
                    continue;
                }
                Thread.Sleep(7000);

                //检测进度退出功能好像还有点问题
                Console.WriteLine($"第{loop}次循环");
                loop++;
                Console.WriteLine(driver.FindElement(By.XPath("//*[@id='rms-studyRate']")).Text);
                if (driver.FindElement(By.XPath("//*[@id='rms-studyRate']")).Text == "100")
                { break; } 
            }
           
        }

        // 当完成目标后进行其他操作或退出程序
        // Do something else or exit the program when the target is achieved
    }
}
