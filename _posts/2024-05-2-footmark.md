---
layout: post
title: 你好
date: 2024-05-2
Author: 李然
categories: 
tags: [sample, document]
comments: true
--- 

在这里简单整理一下以前的内容，方便可以到处查看，如果你碰巧进来了，欢迎留下你的脚印

## mall

商品查看，购物车，下单，支付，退款

## shop

基础信息，营业状态

## market

拼团，满赠，免费零商品

## privilege

给用户发放生日奖励，通过分析用户消费，充值数据给用户发放对应奖励，并通过微信推送订阅消息。员工端查看营销数据


## urlcenter

链接中心，相册，下载

## kds

实现餐厅厨房系统，包含1客房下单 2厨房kds系统显示订单，播报提醒 3厨师制作，出餐 4通知服务员取餐 5送达客户

## im

基于第三方socket实现服务器主动推送消息，实现以下功能：播报系统，通知前端刷新页面，推送消息

## kds

    整体思路：

    1客房下单 2厨房kds系统展示订单，播报提醒 3厨师制作，出餐 4通知服务员取餐 5送达客户

    订单表：catering_order

    订单商品表：catering_order_goods 

    下单
    1：同一个客人前后下单，在同一个订单内

    2：排序号每天3点重置   格式#01

    3：推送im通知。pad订单推送后厨，叫号屏。手牌订单推送收银台



    说明：主要依据catering_order_goods 表 流转状态，查询数据等

    主要操作

    开始制作 ：
    已取消商品无法开始制作

    已制作 幂等

    制作完成
    只有正在制作的商品才可以操作制作完成

    通知取餐 
    包含两个方面

    1流转订单状态(已通知的可再次通知，保证幂等)

    2推送im通知（后续->通过数店通知服务员）

    取餐完成 
    制作完成的才可以取餐

    未通知也可以取餐

    取消订单 
    包含两个方面

    1流转订单状态（只有未开始制作的才可以取消）

    2推送取消事件（当前需求，只刷新后厨展示，不播报）

    清洁房间
    流转订单已完成

付费中心-短信

    注：目前只支持前付费

    购买短信套餐包
    下单购买成功后之后会回调，

    1、短信账户余额表， 增加余额

    2、套餐包队列表， 入队列新套餐记录

    3、每日定时过期套餐包

    扣费大致流程

    https://s2.loli.net/2024/08/02/NmBVx3QHW4DXu2a.png

    扣费
    1、costcenter-rpc(SmsSendPrePay)冻结短信条数

    生成一个订单号，

    冻结短信条数，

    返回单号，可根据单号查询冻结信息

    2、costcenter-rpc(SmsSendPay)扣费

    释放冻结短信条数，

    判断余额是否足够，

    事务里扣余额表、套餐包队列。扣费成功，

    写入短信支付记录表。

    3、costcenter-rpc(SmsReleaseFreezeNum)取消冻结

    写入取消冻结队列，队列消费，消费失败重新写入一条取消冻结记录

    发送结果
    1、costcenter-rpc(SmsSendStatusCallBack)发送

    成功则修改状态为发送成功

    失败则退款，修改状态为发送失败



    2、costcenter-rpc(SmsReceiptStatusCallBack)回执

    写入回执队列，队列消费，消费失败重新写入待回执列表（重试机制，确保成功）



    发送短信
    1、sms-rpc：先冻结短信条数->写入发送记录表->把记录id写入待发送队列（冻结短信之后，如果有错误就调用取消冻结）。

    不同品牌写入不同队列，品牌id作为队列key的一部分，每个品牌的短信发送隔离

    短信类型也作为队列key的一部分，验证码，营销，通知短信隔离

    2、sms-rmq：每个品牌每个类型启用一个消费者，Consume自己的待发送队列->调用SmsSendPay扣费->调用短信服务发送短信->SmsSendStatusCallBack上报发送结果->回写短信记录表


    注：单条短信和多条短信的发送逻辑不一样，

    单条写入单条发送的队列，调用短信服务发送单条的接口

    多条写入多条发送的队列，调用短信服务发送多条的接口

    创蓝回调
    1、写入redis队列待消费，为了保证速度，使用rand.Intn()

    2、SmsSendStatusCallBack上报回执结果

    3、修改发送日志表状态

## im
    整体思路

    一台设备对应一个im账号，设备切房间之类操作时，修改设备位置信息

    收银台为一个门店均为同一个im账号

    推送im消息流程：

    1、业务方调用messagecenter-rpc （SendImCall）

    int64 brandId                             = 1;  // 品牌id
    EnumImCallEvent event                     = 2;  // 事件类型
    repeated EnumImCallChannel Channels       = 3;  // 发送渠道
    map<string, string> callParams            = 4;  // 参数map 详细内容,请看内容
    int64 shopId                              = 5;  // 门店id
    repeated DesignatedAccount designatedList = 6;  // 是否指定账号
    EnumImCallMsgType MsgType                 = 7;  // 消息类型 1播报 2刷新 3通知 4pad呼叫
    int64 msgTemplateId                       = 8;  // 模板id MsgType=4时需要

    2、messagecenter-rpc将要推送内容写表，并写入延时队列

    说明：查询播报配置，判断是否需要发送

    推往多个端会写入多条记录，需要多次推送的也写入多条记录（例如：间隔播报多次）

    3、messagecenter-rmq消费延时队列，组装发送内容，调用txim-rpc推送im消息

    说明：发送前做一些校验，例如检查 端是否需要播报，技师是否已下钟等

    {roomName}房间下单了{goodsName}{num}份，请注意
    参数替换，拼出消息内容，

    801房间下单了水饺2份，请注意

