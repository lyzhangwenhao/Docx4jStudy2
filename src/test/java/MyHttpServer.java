/**
 * ClassName: MyHttpServer
 * Description:
 *
 * @author 张文豪
 * @date 2020/10/13 9:00
 */

import com.sun.net.httpserver.Headers;
import com.sun.net.httpserver.HttpExchange;
import com.sun.net.httpserver.HttpHandler;
import com.sun.net.httpserver.HttpServer;

import java.io.IOException;
import java.io.OutputStream;
import java.net.InetSocketAddress;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * 根据Java提供的API实现Http服务器
 */
public class MyHttpServer {

    /**
     * @param args
     * @throws IOException
     */
    public static void main(String[] args) throws IOException {
        //创建HttpServer服务器
        HttpServer httpServer = HttpServer.create(new InetSocketAddress(8080), 10);
        //将 /jay请求交给MyHandler处理器处理
        httpServer.createContext("/", new MyHandler());
        httpServer.start();
    }
}

class MyHandler implements HttpHandler {

    public void handle(HttpExchange httpExchange) throws IOException {
        //请求头
        Headers headers = httpExchange.getRequestHeaders();
        Set<Map.Entry<String, List<String>>> entries = headers.entrySet();

        StringBuffer response = new StringBuffer();
        for (Map.Entry<String, List<String>> entry : entries){
            response.append(entry.toString() + "\n");
        }
        //设置响应头属性及响应信息的长度
        httpExchange.sendResponseHeaders(200, response.length());
        //获得输出流
        OutputStream os = httpExchange.getResponseBody();
        os.write(response.toString().getBytes());
        os.close();
    }
}