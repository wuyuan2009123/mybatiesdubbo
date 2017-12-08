/**   
 * @Project: IFeidaiSer 
 * @Title: HttpUtil.java 
 * @Package com.yxtec.ifeidai.utils.fdfs 
 * @Description: Http请求对象 
 * @author liaogd 
 * @date 2015年3月14日 下午4:33:28 
 * @Copyright: 2015 年 研信科技. All rights reserved  
 * @version V1.0   
 */
package com.springboot.dubbo.util;

/** 
 * @ClassName HttpUtil  
 * @Description Http请求对象 
 * @author liaogd 
 * @date 2015年3月14日  
 *   
 */
import java.io.BufferedOutputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.net.URISyntaxException;
import java.nio.charset.Charset;
import java.security.KeyStore;
import java.security.cert.CertificateException;
import java.security.cert.X509Certificate;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;

import javax.net.ssl.SSLContext;

import org.apache.commons.io.IOUtils;
import org.apache.commons.lang.ArrayUtils;
import org.apache.commons.lang.StringUtils;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.http.HttpException;
import org.apache.http.NameValuePair;
import org.apache.http.ParseException;
import org.apache.http.client.ClientProtocolException;
import org.apache.http.client.entity.UrlEncodedFormEntity;
import org.apache.http.client.methods.CloseableHttpResponse;
import org.apache.http.client.methods.HttpGet;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.client.utils.URIBuilder;
import org.apache.http.config.ConnectionConfig;
import org.apache.http.config.Registry;
import org.apache.http.config.RegistryBuilder;
import org.apache.http.config.SocketConfig;
import org.apache.http.conn.socket.ConnectionSocketFactory;
import org.apache.http.conn.socket.LayeredConnectionSocketFactory;
import org.apache.http.conn.socket.PlainConnectionSocketFactory;
import org.apache.http.conn.ssl.SSLConnectionSocketFactory;
import org.apache.http.conn.ssl.SSLContexts;
import org.apache.http.conn.ssl.TrustStrategy;
import org.apache.http.cookie.Cookie;
import org.apache.http.cookie.CookieOrigin;
import org.apache.http.entity.StringEntity;
import org.apache.http.entity.mime.FormBodyPart;
import org.apache.http.entity.mime.HttpMultipartMode;
import org.apache.http.entity.mime.MultipartEntityBuilder;
import org.apache.http.impl.client.BasicCookieStore;
import org.apache.http.impl.client.CloseableHttpClient;
import org.apache.http.impl.client.HttpClientBuilder;
import org.apache.http.impl.conn.PoolingHttpClientConnectionManager;
import org.apache.http.impl.cookie.BasicClientCookie;
import org.apache.http.impl.cookie.BestMatchSpec;
import org.apache.http.message.BasicNameValuePair;
import org.apache.http.util.EntityUtils;


public class HttpUtil {
    private static final Log LOG = LogFactory.getLog(HttpUtil.class);
    private static int bufferSize = 1024;

    private CloseableHttpClient client;
    private BasicCookieStore cookieStore;
    private static final String DEFAULTENCODING = "UTF-8";

    private HttpUtil(boolean isSSL) {
        ConnectionConfig connConfig;
        SocketConfig socketConfig;
        ConnectionSocketFactory plainSF;
        KeyStore trustStore;
        SSLContext sslContext;
        LayeredConnectionSocketFactory sslSF;
        Registry<ConnectionSocketFactory> registry;
        PoolingHttpClientConnectionManager connManager;
        BasicCookieStore cookieStore;
        // 设置连接参数
        connConfig = ConnectionConfig.custom().setCharset(Charset.forName(DEFAULTENCODING)).build();
        socketConfig = SocketConfig.custom().setSoTimeout(100000).build();
        RegistryBuilder<ConnectionSocketFactory> registryBuilder = RegistryBuilder
                .<ConnectionSocketFactory> create();
        plainSF = new PlainConnectionSocketFactory();
        registryBuilder.register("http", plainSF);
        if (isSSL) {
            // 指定信任密钥存储对象和连接套接字工厂
            try {
                trustStore = KeyStore.getInstance(KeyStore.getDefaultType());
                sslContext = SSLContexts.custom().useTLS()
                        .loadTrustMaterial(trustStore, new AnyTrustStrategy()).build();
                sslSF = new SSLConnectionSocketFactory(sslContext,
                        SSLConnectionSocketFactory.ALLOW_ALL_HOSTNAME_VERIFIER);
                registryBuilder.register("https", sslSF);
            } catch (Exception e) {
                throw new RuntimeException(e);
            }
        }
        registry = registryBuilder.build();
        // 设置连接管理
        connManager = new PoolingHttpClientConnectionManager(registry);
        connManager.setDefaultConnectionConfig(connConfig);
        connManager.setDefaultSocketConfig(socketConfig);

        // 指定cookie存储对象
        cookieStore = new BasicCookieStore();
        // 构建客户
        client = HttpClientBuilder.create().setDefaultCookieStore(cookieStore)
                .setConnectionManager(connManager).build();
    }

    public static HttpUtil getInstance(boolean isSSL) {
        return new HttpUtil(isSSL);
    }

    /** 
     * @Title: close 
     * @Description: 关闭HttpClient，释放资源
     * @return void    返回类型
     */
    public void close() {
        IOUtils.closeQuietly(client);
    }

    /** 
     * @Title: doGetReturnString 
     * @param url
     * @param queryParams
     * @param charset
     * @return
     * @throws ParseException
     * @throws IOException 参数说明
     * @return String    返回类型
     */
    public String doGetReturnString(String url, Map<String, Object> queryParams, String charset)
            throws ParseException, IOException {
        CloseableHttpResponse response = null;
        String responseStr = null;
        try {
            response = doGet(url, queryParams);
            responseStr = EntityUtils.toString(response.getEntity(), charset);

            return responseStr;
        } finally {
            IOUtils.closeQuietly(response);
            IOUtils.closeQuietly(client);
        }
    }

    /**
     * 基本的Get请求
     * @param url 请求url
     * @param queryParams 请求头的查询参数
     * @return
     * @throws URISyntaxException 
     * @throws IOException 
     * @throws ClientProtocolException 
     */
    public CloseableHttpResponse doGet(String url, Map<String, Object> queryParams) {
        try {
            HttpGet gm = new HttpGet();
            URIBuilder builder = new URIBuilder(url);
            // 填入查询参数
            if (queryParams != null && !queryParams.isEmpty()) {
                builder.setParameters(HttpUtil.paramsConverter(queryParams));
            }
            gm.setURI(builder.build());

            // 由服务器端自动关闭链接
            gm.setHeader("Connection", "close");
            int code = 0;
            CloseableHttpResponse closeableHttpResponse = client.execute(gm);
            code = closeableHttpResponse.getStatusLine().getStatusCode();
            return closeableHttpResponse;
        } catch (Exception e) {
            LOG.error("调用" + url + "出错！ 参数：" + queryParams, e);
        }
        return null;
    }

    public CloseableHttpResponse doGet(String url, Map<String, Object> queryParams,
            Map<String, String> headMap) {
        try {
            int code = 0;
            HttpGet gm = new HttpGet();
            URIBuilder builder = new URIBuilder(url);
            // 填入查询参数
            if (queryParams != null && !queryParams.isEmpty()) {
                // 将设置新的参数修改为添加新的参数，防止之前有参数带过来
                builder.addParameters(HttpUtil.paramsConverter(queryParams));
                // builder.setParameters(HttpUtil.paramsConverter(queryParams));
            }
            gm = setHead(gm, headMap);
            gm.setURI(builder.build());
            // 由服务器端自动关闭链接
            gm.setHeader("Connection", "close");
            
            CloseableHttpResponse closeableHttpResponse = client.execute(gm);
            code = closeableHttpResponse.getStatusLine().getStatusCode();
            return closeableHttpResponse;
        } catch (Exception e) {
            LOG.error("调用" + url + "出错！ 参数：" + queryParams + " " + headMap, e);
        }
        return null;
    }

    /**
     * 设置请求头
     * @Title: setHead 
     * @Description: 设置请求头
     * @param gm
     * @param headMap
     * @return 参数说明
     * @return HttpGet    返回类型
     */
    public HttpGet setHead(HttpGet gm, Map<String, String> headMap) {

        Set<Entry<String, String>> paramsSet = headMap.entrySet();
        for (Entry<String, String> paramEntry : paramsSet) {
            gm.setHeader(paramEntry.getKey(), paramEntry.getValue());
        }
        return gm;
    }

    public HttpPost setHead(HttpPost pm, Map<String, String> headMap) {

        Set<Entry<String, String>> paramsSet = headMap.entrySet();
        for (Entry<String, String> paramEntry : paramsSet) {
            pm.setHeader(paramEntry.getKey(), paramEntry.getValue());
        }
        return pm;
    }

    /** 
     * @Title: doPostReturnString 
     * @Description: 基本的Post请求，直接返回结果。
     * @param url
     * @param queryParams
     * @param formParams
     * @param charset
     * @return
     * @throws ParseException
     * @throws IOException 参数说明
     * @return String    返回类型
     */
    public String doPostReturnString(String url, Map<String, Object> queryParams,
            Map<String, Object> formParams, String charset) throws ParseException, IOException {
        CloseableHttpResponse response = null;
        String responseStr = null;
        try {
            response = doPost(url, queryParams, formParams);
            responseStr = EntityUtils.toString(response.getEntity(), charset);

            return responseStr;
        } finally {
            IOUtils.closeQuietly(response);
            IOUtils.closeQuietly(client);
        }
    }

    /** 
     * @Title: doPostReturnString 
     * @Description: 基本的Post请求，直接返回结果。
     * @param url
     * @param queryParams
     * @param formParams
     * @return
     * @throws ParseException
     * @throws IOException 参数说明
     * @return String    返回类型
     */
    public String doPostReturnString(String url, Map<String, Object> queryParams,
            Map<String, Object> formParams) throws ParseException, IOException {
        return doPostReturnString(url, queryParams, formParams, null);
    }

    /** 
     * @Title: doPostWithoutReturn 
     * @Description: 调用URL请求不用返回结果。
     * @param url
     * @param queryParams
     * @param formParams 参数说明
     * @return void    返回类型
     */
    public void doPostWithoutReturn(String url, Map<String, Object> queryParams,
            Map<String, Object> formParams) {
        CloseableHttpResponse response = null;
        try {
            response = doPost(url, queryParams, formParams);
        } finally {
            IOUtils.closeQuietly(response);
            IOUtils.closeQuietly(client);
        }
    }

    /**
     * 基本的Post请求
     * @param url 请求url
     * @param queryParams 请求头的查询参数
     * @param formParams post表单的参敄
     * @return
     * @throws URISyntaxException 
     * @throws IOException 
     * @throws ClientProtocolException 
     */
    private CloseableHttpResponse doPost(String url, Map<String, Object> queryParams,
            Map<String, Object> formParams) {
        try {
            HttpPost pm = new HttpPost();
            URIBuilder builder = new URIBuilder(url);
            // 填入查询参数
            if (queryParams != null && !queryParams.isEmpty()) {
                builder.setParameters(HttpUtil.paramsConverter(queryParams));
            }
            pm.setURI(builder.build());
            // 填入表单参数
            if (formParams != null && !formParams.isEmpty()) {
                pm.setEntity(new UrlEncodedFormEntity(HttpUtil.paramsConverter(formParams), "UTF-8"));
            }
            // 由服务器端自动关闭链接
            pm.setHeader("Connection", "close");
            int code = 0;
            CloseableHttpResponse closeableHttpResponse = client.execute(pm);
            code = closeableHttpResponse.getStatusLine().getStatusCode();
            return closeableHttpResponse;
        } catch (Exception e) {
            LOG.error("调用" + url + "出错！ 参数：" + queryParams + " " + formParams, e);
        }
        return null;
    }

    public CloseableHttpResponse doPost(String url, Map<String, Object> queryParams,
            Map<String, Object> formParams, Map<String, String> headMap, String jsonOrXml,
            String reqEncoding) {
        try {
            HttpPost pm = new HttpPost();
            URIBuilder builder = new URIBuilder(url);

            // 填入查询参数
            if (queryParams != null && !queryParams.isEmpty()) {
                builder.setParameters(HttpUtil.paramsConverter(queryParams));
            }
            pm.setURI(builder.build());

            pm = setHead(pm, headMap);

            if (StringUtils.isNotBlank(jsonOrXml)) {
                if (headMap.isEmpty()) {
                    pm.setHeader("Content-Type", "application/json");
                }
                pm.setEntity(new StringEntity(jsonOrXml, reqEncoding));
            } else {
                // 填入表单参数
                if (formParams != null && !formParams.isEmpty()) {
                    pm.setEntity(new UrlEncodedFormEntity(HttpUtil.paramsConverter(formParams),
                            reqEncoding));
                }
            }
            // 由服务器端自动关闭链接
            pm.setHeader("Connection", "close");
            int code = 0;
            
            CloseableHttpResponse closeableHttpResponse = client.execute(pm);
            code = closeableHttpResponse.getStatusLine().getStatusCode();
            
            return closeableHttpResponse;
        } catch (Exception e) {
            
            LOG.error("调用" + url + "出错！ 参数：" + queryParams + " " + formParams + " " + jsonOrXml
                    + " " + reqEncoding, e);
        }
        return null;
    }

    /**
     * 多块Post请求
     * @param url 请求url
     * @param queryParams 请求头的查询参数
     * @param formParts post表单的参敄,支持字符丄-文件(FilePart)和字符串-字符丄(StringPart)形式的参敄
     * @param maxCount 多尝试请求的次数
     * @return
     * @throws URISyntaxException 
     * @throws ClientProtocolException 
     * @throws HttpException
     * @throws IOException
     */
    public CloseableHttpResponse multipartPost(String url, Map<String, Object> queryParams,
            List<FormBodyPart> formParts) {
        try {
            HttpPost pm = new HttpPost();
            URIBuilder builder = new URIBuilder(url);
            // 填入查询参数
            if (queryParams != null && !queryParams.isEmpty()) {
                builder.setParameters(HttpUtil.paramsConverter(queryParams));
            }
            pm.setURI(builder.build());
            // 填入表单参数
            if (formParts != null && !formParts.isEmpty()) {
                MultipartEntityBuilder entityBuilder = MultipartEntityBuilder.create();
                entityBuilder = entityBuilder.setMode(HttpMultipartMode.BROWSER_COMPATIBLE);
                for (FormBodyPart formPart : formParts) {
                    entityBuilder = entityBuilder.addPart(formPart.getName(), formPart.getBody());
                }
                pm.setEntity(entityBuilder.build());
            }
            // 由服务器端自动关闭链接
            pm.setHeader("Connection", "close");
            int code = 0;
            CloseableHttpResponse closeableHttpResponse = client.execute(pm);
            code = closeableHttpResponse.getStatusLine().getStatusCode();
            return closeableHttpResponse;
        } catch (Exception e) {
            LOG.error("调用" + url + "出错！ 参数：" + queryParams, e);
        }
        return null;
    }

    /**
     * 获取当前Http客户端状态中的Cookie
     * @param domain 作用埄
     * @param port 端口 传null 默认80
     * @param path Cookie路径 传null 默认"/"
     * @param useSecure Cookie是否采用安全机制 传null 默认false
     * @return
     */
    public Map<String, Cookie> getCookie(String domain, Integer port, String path, Boolean useSecure) {
        if (domain == null) {
            return null;
        }
        if (port == null) {
            port = 80;
        }
        if (path == null) {
            path = "/";
        }
        if (useSecure == null) {
            useSecure = false;
        }
        List<Cookie> cookies = cookieStore.getCookies();
        if (cookies == null || cookies.isEmpty()) {
            return null;
        }

        CookieOrigin origin = new CookieOrigin(domain, port, path, useSecure);
        BestMatchSpec cookieSpec = new BestMatchSpec();
        Map<String, Cookie> retVal = new HashMap<>();
        for (Cookie cookie : cookies) {
            if (cookieSpec.match(cookie, origin)) {
                retVal.put(cookie.getName(), cookie);
            }
        }
        return retVal;
    }

    /**
     * 批量设置Cookie
     * @param cookies cookie
     * @param domain 作用埄  不可为空
     * @param path 路径 传null默认丄 "/"
     * @param useSecure 是否使用安全机制 传null 默认为false
     * @return 是否成功设置cookie
     */
    public boolean setCookie(Map<String, String> cookies, String domain, String path,
            Boolean useSecure) {
        synchronized (cookieStore) {
            if (domain == null) {
                return false;
            }
            if (path == null) {
                path = "/";
            }
            if (useSecure == null) {
                useSecure = false;
            }
            if (cookies == null || cookies.isEmpty()) {
                return true;
            }
            Set<Entry<String, String>> set = cookies.entrySet();
            String key = null;
            String value = null;
            for (Entry<String, String> entry : set) {
                key = entry.getKey();
                if (key == null || key.isEmpty() || value == null || value.isEmpty()) {
                    throw new IllegalArgumentException(
                            "cookies key and value both can not be empty");
                }
                BasicClientCookie cookie = new BasicClientCookie(key, value);
                cookie.setDomain(domain);
                cookie.setPath(path);
                cookie.setSecure(useSecure);
                cookieStore.addCookie(cookie);
            }
            return true;
        }
    }

    /**
     * 设置单个Cookie
     * @param key Cookie
     * @param value Cookie
     * @param domain 作用埄  不可为空
     * @param path 路径 传null默认丄 "/"
     * @param useSecure 是否使用安全机制 传null 默认为false
     * @return 是否成功设置cookie
     */
    public boolean setCookie(String key, String value, String domain, String path, Boolean useSecure) {
        Map<String, String> cookies = new HashMap<>();
        cookies.put(key, value);
        return setCookie(cookies, domain, path, useSecure);
    }

    /**
     * 设置post请求表单参数
     * @param params 参数map集合
     */
    private static List<NameValuePair> paramsConverter(Map<String, Object> params) {
        List<NameValuePair> nvps = new LinkedList<>();
        Set<Entry<String, Object>> paramsSet = params.entrySet();

        for (Entry<String, Object> paramEntry : paramsSet) {
            nvps.add(new BasicNameValuePair(paramEntry.getKey(), (String) paramEntry.getValue()));
        }
        return nvps;
    }

    /**
     * 将流读取成字节数组
     * @param in 输入浄
     */
    public static byte[] readToByte(InputStream in) {
        if (in == null) {
            return ArrayUtils.EMPTY_BYTE_ARRAY;
        }
        try {
            ByteArrayOutputStream out = new ByteArrayOutputStream(bufferSize);
            byte[] temp = new byte[bufferSize];
            int size = 0;
            while ((size = in.read(temp)) != -1) {
                out.write(temp, 0, size);
            }
            return out.toByteArray();
        } catch (Exception e) {
            LOG.error("读取返回内容出错", e);
        } finally {
            IOUtils.closeQuietly(in);
        }
        return ArrayUtils.EMPTY_BYTE_ARRAY;
    }

    /**
     * 将流写入文件
     * @param in 输入浄
     */
    public static void writeToFile(InputStream in, String localFilePath) {
        if (in == null) {
            return;
        }
        try {
            File file = new File(localFilePath);
            String localpath = "";
            if (localFilePath.lastIndexOf('/') > 0) {
                localpath = localFilePath.substring(0, localFilePath.lastIndexOf('/'));
            } else if (localFilePath.lastIndexOf('\\') > 0) {
                localpath = localFilePath.substring(0, localFilePath.lastIndexOf('\\'));
            }
            File f = new File(localpath);
            // 文件夹不存在时创建文件夹
            if (!f.exists()) {
                f.mkdirs();
            }
            BufferedOutputStream out = new BufferedOutputStream(new FileOutputStream(file));
            byte[] buf = new byte[2048];
            int length = in.read(buf);
            while (length != -1) {
                out.write(buf, 0, length);
                length = in.read(buf);
            }
            out.close();
        } catch (Exception e) {
            LOG.error("读取返回内容出错", e);
        } finally {
            IOUtils.closeQuietly(in);
        }
    }

    /**
     * 将流转成指定编码的字符串
     * @param in 输入浄
     * @param encoding 编码格式
     */
    public static String readStream(InputStream in, String encoding) {
        if (in == null) {
            return null;
        }
        try {
            InputStreamReader inReader = null;
            if (encoding == null) {
                inReader = new InputStreamReader(in, DEFAULTENCODING);
            } else {
                inReader = new InputStreamReader(in, encoding);
            }
            char[] buffer = new char[bufferSize];
            int readLen = 0;
            StringBuilder sb = new StringBuilder();
            while ((readLen = inReader.read(buffer)) != -1) {
                sb.append(buffer, 0, readLen);
            }
            inReader.close();
            return sb.toString();
        } catch (Exception e) {
            LOG.error("读取返回内容出错", e);
        } finally {
            IOUtils.closeQuietly(in);
        }
        return null;
    }

    class AnyTrustStrategy implements TrustStrategy {
    	@Override
        public boolean isTrusted(X509Certificate[] chain, String authType)
                throws CertificateException {
            return true;
        }

    }
}
