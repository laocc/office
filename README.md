# 安装：

composer:

```
composer require laocc/session
```

# 引用方法：

```php
<?php

$config = [
    'drive' => 'redis',//驱动方式：redis,file
    'key' => 'PHPSESSID',//session在cookies中的名称
    'delay' => 0,//是否自动延期
    'prefix' => '',//session保存在redis或file中的键名前缀
    'path' => '/',//设定会话 cookie 的路径，一般就为网站根目录，也可以指定如：/admin
    'domain' => 'domain',//在哪个域名下有效，domain则取值实际请求的域名，或在这里指定域名
    'limiter' => 'nocache',//客户端缓存方法，没太明白这个
    'expire' => 86400,//session保存时间，在redis中过了这个时间就删除，file下无作用
    'cookie' => 86400,//客户端cookies保存时间
    
    //redis的连接方式，若能确保在创建对象后setRedis，可以不用填这个
    'redis' => [
        'host' => '127.0.0.1',
        'port' => 6379,
        'db' => 0,
        'password' => '',
    ]
];
$_session = new Session($config);


//若采用的是redis，且在这之前其他地方创建过redis连接，这里可以先送入redis实例，省略session再次连redis
//须注意：redis表若与之前的连接实例不同，这里不要送入，否则会保存到送入的实例表中

//$redis = new \Redis(...);

$_session->setRedis($redis);

```

# 数据说明：

若某页面原则上是不会改变任何session，为保险起见，可在页面任何地方加：session_abort();用于丢弃当前进程所有对session的改动；

本类只是改变PHP存取session的介质，在使用方面没有影响，如：`$_SESSION['v']=123`，`$v=$_SESSION['v']`；

本插件实现用redis保存session，且每个session的生存期从其自身被定义时计算起，而非PHP本身统一设置

有一个问题须注意：`$_SESSION['name']=abc`；之后若再次给`$_SESSION['name']`赋其他不同的值，则其生存期以第二次赋值起算起 但是，若第二次赋值与之前的值相同，并不会改变其生存期

### 简单设置

如果只是想存到redis也可以直接设置，或修改`php.ini`:

```
ini_set('session.save_handler', 'redis');
ini_set('session.save_path', 'tcp://127.0.0.1:6379');
ini_set('session.save_path', '/tmp/redis.sock?database=0');
```

### file查看session

php.ini中默认保存到PHP，也就是服务器某个目录中， 比如默认：`session.save_path = "/tmp"`

在`/tmp`中所有`[sess_****]`文件即为session内容

### redis中查看session

如果指定redis作为介质，则用下列方法可查看session内容

```
[root@localhost ~]# redis-cli
127.0.0.1:6379> ping
PONG
```

列出所有键：

```
127.0.0.1:6379> keys PHPREDIS*
1) "PHPREDIS_SESSION :57105pkee2ov7b49il470ctv51"
```

显示内容：

```
127.0.0.1:6379> get PHPREDIS_SESSION :57105pkee2ov7b49il470ctv51
"val|s:19:\"2017-09-16 14:47:45\";"
```