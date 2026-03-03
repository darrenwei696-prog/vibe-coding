// 信息图标组件
function InfoIcon(props: React.SVGProps<SVGSVGElement>) {
  return (
    <svg
      {...props}
      xmlns="http://www.w3.org/2000/svg"
      viewBox="0 0 20 20"
      fill="currentColor"
    >
      <path
        fillRule="evenodd"
        d="M18 10a8 8 0 11-16 0 8 8 0 0116 0zm-7-4a1 1 0 11-2 0 1 1 0 012 0zM9 9a1 1 0 000 2v3a1 1 0 001 1h1a1 1 0 100-2v-3a1 1 0 00-1-1H9z"
        clipRule="evenodd"
      />
    </svg>
  );
}

import { useState, useRef, useEffect } from 'react';
import { motion } from 'framer-motion';
import { toast } from 'sonner';
import { 
  Mail, 
  Upload, 
  FileSpreadsheet, 
  FileText, 
  Send, 
  ChevronRight, 
  CheckCircle, 
  AlertCircle,
  RotateCcw,
  ChevronDown,
  ChevronUp
} from 'lucide-react';
import { cn } from '@/lib/utils';

interface Recipient {
  name: string;
  email: string;
}

interface ProcessingResult {
  success: boolean;
  recipient: Recipient;
  message?: string;
  subject?: string;
}

export default function EmailAssistant() {
  // 状态管理
  const [senderEmail, setSenderEmail] = useState('');
  const [recipientsFile, setRecipientsFile] = useState<File | null>(null);
  const [emailContentFile, setEmailContentFile] = useState<File | null>(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [results, setResults] = useState<ProcessingResult[]>([]);
  const [showResults, setShowResults] = useState(false);
  // SMTP设置状态
  const [showSmtpSettings, setShowSmtpSettings] = useState(false);
  const [smtpServer, setSmtpServer] = useState('');
  const [smtpPort, setSmtpPort] = useState('');
  const [smtpPassword, setSmtpPassword] = useState('');
  const [showPassword, setShowPassword] = useState(false);
  // 邮件主题状态
  const [emailSubject, setEmailSubject] = useState('');
  const [detectedSubject, setDetectedSubject] = useState('');
  
  // 文件输入引用
  const recipientsFileInputRef = useRef<HTMLInputElement>(null);
  const emailContentFileInputRef = useRef<HTMLInputElement>(null);

  // 为了方便测试，可以添加一些示例数据填充功能
  const fillWithTestData = () => {
    setSenderEmail('test@example.com');
    setEmailSubject('测试邮件主题');
    setShowSmtpSettings(true);
    setSmtpServer('smtp.example.com');
    setSmtpPort('587');
    
    // 显示提示
    toast.info('已填充测试数据，请上传文件后开始处理');
  };

  // 处理收件人文件上传
  const handleRecipientsFileUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (file) {
      if (file.type !== 'application/vnd.ms-excel' && file.type !== 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') {
        toast.error('请上传Excel文件(.xls或.xlsx)');
        return;
      }
      setRecipientsFile(file);
      toast.success(`已上传：${file.name}`);
    }
  };

  // 处理邮件内容文件上传
  const handleEmailContentFileUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (file) {
      if (file.type !== 'application/msword' && file.type !== 'application/vnd.openxmlformats-officedocument.wordprocessingml.document') {
        toast.error('请上传Word文档(.doc或.docx)');
        return;
      }
      setEmailContentFile(file);
      toast.success(`已上传：${file.name}`);
    }
  };

  // 验证邮箱格式
  const validateEmail = (email: string): boolean => {
    const re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return re.test(email);
  };

  // 获取邮箱提供商的默认SMTP设置
  const getDefaultSmtpSettings = (email: string) => {
    if (email.includes('@gmail.com')) {
      setSmtpServer('smtp.gmail.com');
      setSmtpPort('587');
    } else if (email.includes('@outlook.com') || email.includes('@hotmail.com')) {
      setSmtpServer('smtp.office365.com');
      setSmtpPort('587');
    } else if (email.includes('@qq.com')) {
      setSmtpServer('smtp.qq.com');
      setSmtpPort('465');
    } else if (email.includes('@163.com')) {
      setSmtpServer('smtp.163.com');
      setSmtpPort('465');
    } else if (email.includes('@126.com')) {
      setSmtpServer('smtp.126.com');
      setSmtpPort('465');
    } else if (email.includes('@sina.com')) {
      setSmtpServer('smtp.sina.com');
      setSmtpPort('465');
    }
  };

  // 处理发件人邮箱变化，自动填充默认SMTP设置
  useEffect(() => {
    if (validateEmail(senderEmail)) {
      getDefaultSmtpSettings(senderEmail);
    }
  }, [senderEmail]);

  // 处理表单提交
  const handleSubmit = async (event: React.FormEvent) => {
    event.preventDefault();
    
    // 表单验证
    if (!validateEmail(senderEmail)) {
      toast.error('请输入有效的发件人邮箱');
      return;
    }
    
    if (!recipientsFile) {
      toast.error('请上传收件人表格');
      return;
    }
    
    if (!emailContentFile) {
      toast.error('请上传邮件内容文档');
      return;
    }
    
    // 开始处理
    setIsProcessing(true);
    setShowResults(false);
    
    // 模拟处理过程
    try {
      // 模拟从Excel解析收件人数据
      // 在实际应用中，这里需要使用如xlsx等库解析Excel文件
      // 改进后的模拟数据生成逻辑，基于文件名和时间生成更真实的模拟数据
      const generateMockRecipients = (fileName: string): Recipient[] => {
        // 使用文件名和时间戳生成伪随机数种子
        const seed = fileName.split('').reduce((acc, char) => acc + char.charCodeAt(0), 0) + Date.now();
        const random = (max: number) => {
          // 简单的线性同余生成器
          return ((seed * 1103515245 + 12345) % 2147483647) % max;
        };
        
        // 常见姓氏和名字
        const lastName = ['张', '王', '李', '赵', '刘', '陈', '杨', '黄', '周', '吴'];
        const firstName = ['伟', '芳', '娜', '秀英', '敏', '静', '强', '磊', '军', '洋'];
        
        // 生成3-6个随机收件人
        const count = 3 + random(4);
        const recipients: Recipient[] = [];
        
        for (let i = 0; i < count; i++) {
          const ln = lastName[random(lastName.length)];
          const fn = firstName[random(firstName.length)];
          const domain = ['gmail.com', 'outlook.com', 'qq.com', '163.com', '126.com'][random(5)];
          // 生成类似真实邮箱的格式
          const email = `${ln}${fn}${i+1}@${domain}`.toLowerCase();
          
          recipients.push({
            name: ln + fn,
            email
          });
        }
        
        return recipients;
      };
      
      const mockRecipients = generateMockRecipients(recipientsFile.name);

      // 模拟从Word文档提取红色背景内容作为邮件主题
      // 在实际应用中，这里需要使用如docx等库解析Word文档并识别红色背景内容
      const mockDetectedSubjects = [
        "重要通知：系统升级维护",
        "项目进度更新与下一步计划",
        "会议纪要与行动项分配",
        "季度报告与业绩分析",
        "团队活动安排与报名通知"
      ];
      const detectedSubject = mockDetectedSubjects[Math.floor(Math.random() * mockDetectedSubjects.length)];
      setDetectedSubject(detectedSubject);
      setEmailSubject(detectedSubject);
      
      // 模拟处理每个收件人
      const processingResults: ProcessingResult[] = await Promise.all(
        mockRecipients.map(async (recipient) => {
          // 模拟处理时间
          await new Promise(resolve => setTimeout(resolve, 500 + Math.random() * 1000));
          
          // 模拟成功/失败结果
          // 在实际应用中，这里会使用真实的SMTP服务发送邮件
          let success = Math.random() > 0.2; // 80%成功率
          let message = success ? '邮件发送成功' : '邮件发送失败';
          
          // 模拟SMTP相关的失败情况
          if (!success) {
            const errorTypes = [
              'SMTP认证失败，请检查授权码是否正确',
              'SMTP服务器连接超时，请检查网络或服务器设置',
              '邮箱未开启SMTP服务，请在邮箱设置中开启',
              '收件人邮箱不存在或无效'
            ];
            message = errorTypes[Math.floor(Math.random() * errorTypes.length)];
          }
          
          return {
            success,
            recipient,
            message,
            subject: detectedSubject
          };
        })
      );
      
      setResults(processingResults);
      setShowResults(true);
      
      // 滚动到结果区域
      setTimeout(() => {
        const resultsSection = document.getElementById('results-section');
        if (resultsSection) {
          resultsSection.scrollIntoView({ behavior: 'smooth' });
        }
      }, 100);
      
    } catch (error) {
      toast.error('处理过程中发生错误');
      console.error('处理错误:', error);
    } finally {
      setIsProcessing(false);
    }
  };

  // 重置表单
  const resetForm = () => {
    setSenderEmail('');
    setRecipientsFile(null);
    setEmailContentFile(null);
    setResults([]);
    setShowResults(false);
    setShowSmtpSettings(false);
    setSmtpServer('');
    setSmtpPort('');
    setSmtpPassword('');
    setShowPassword(false);
    setEmailSubject('');
    setDetectedSubject('');
    
    // 重置文件输入
    if (recipientsFileInputRef.current) {
      recipientsFileInputRef.current.value = '';
    }
    if (emailContentFileInputRef.current) {
      emailContentFileInputRef.current.value = '';
    }
  };

  // 动画变量
  const containerVariants = {
    hidden: { opacity: 0 },
    visible: { 
      opacity: 1,
      transition: { 
        staggerChildren: 0.1,
        delayChildren: 0.2
      }
    }
  };
  
  const itemVariants = {
    hidden: { y: 20, opacity: 0 },
    visible: { 
      y: 0, 
      opacity: 1,
      transition: { duration: 0.5, ease: "easeOut" }
    }
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-50 to-indigo-50 dark:from-gray-900 dark:to-blue-950 py-12 px-4 sm:px-6 lg:px-8 font-sans">
      <div className="max-w-4xl mx-auto">
        {/* 标题部分 */}
        <motion.div 
          initial={{ y: -20, opacity: 0 }}
          animate={{ y: 0, opacity: 1 }}
          transition={{ duration: 0.5 }}
          className="text-center mb-12"
        >
          <h1 className="text-4xl font-bold text-gray-900 dark:text-white mb-2 flex items-center justify-center">
            <Mail className="mr-2 text-blue-600 dark:text-blue-400" size={32} />
            邮件助理 <span className="text-blue-600 dark:text-blue-400">(comcom)</span>
          </h1>
          <p className="text-gray-600 dark:text-gray-300 text-lg max-w-2xl mx-auto">
            轻松批量处理邮件，自动替换收件人信息，提高工作效率
          </p>
        </motion.div>

        {/* 表单卡片 */}
        <motion.div 
          className="bg-white dark:bg-gray-800 rounded-xl shadow-xl overflow-hidden mb-12 border border-gray-100 dark:border-gray-700"
          initial={{ y: 20, opacity: 0 }}
          animate={{ y: 0, opacity: 1 }}
          transition={{ duration: 0.5, delay: 0.2 }}
        >
          <div className="p-8">
            <form onSubmit={handleSubmit}>
              <motion.div 
                variants={containerVariants}
                initial="hidden"
                animate="visible"
                className="space-y-8"
              >
                {/* 步骤指示器 */}
                <motion.div variants={itemVariants} className="flex items-center justify-center mb-8">
                  <div className="flex items-center w-full max-w-md">
                    {/* 步骤1 */}
                    <div className="flex flex-col items-center">
                      <div className="w-10 h-10 rounded-full bg-blue-600 flex items-center justify-center text-white font-bold">1</div>
                      <span className="text-xs mt-2 text-gray-600 dark:text-gray-300">填写发件邮箱</span>
                    </div>
                    
                    {/* 连接线 */}
                    <div className="flex-1 h-1 bg-blue-200 dark:bg-blue-800 mx-4"></div>
                    
                    {/* 步骤2 */}
                    <div className="flex flex-col items-center">
                      <div className="w-10 h-10 rounded-full bg-blue-600 flex items-center justify-center text-white font-bold">2</div>
                      <span className="text-xs mt-2 text-gray-600 dark:text-gray-300">上传收件人表格</span>
                    </div>
                    
                    {/* 连接线 */}
                    <div className="flex-1 h-1 bg-blue-200 dark:bg-blue-800 mx-4"></div>
                    
                    {/* 步骤3 */}
                    <div className="flex flex-col items-center">
                      <div className="w-10 h-10 rounded-full bg-blue-600 flex items-center justify-center text-white font-bold">3</div>
                      <span className="text-xs mt-2 text-gray-600 dark:text-gray-300">上传邮件内容</span>
                    </div>
                  </div>
                </motion.div>

                {/* 发件人邮箱输入 */}
                <motion.div variants={itemVariants} className="space-y-2">
                  <label htmlFor="senderEmail" className="block text-sm font-medium text-gray-700 dark:text-gray-300">
                    发件人邮箱
                  </label>
                  <div className="relative">
                    <div className="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none">
                      <Mail className="h-5 w-5 text-gray-400" />
                    </div>
                    <input
                      type="email"
                      id="senderEmail"
                      value={senderEmail}
                      onChange={(e) => setSenderEmail(e.target.value)}
                      className="block w-full pl-10 pr-3 py-3 border border-gray-300 dark:border-gray-600 rounded-lg bg-white dark:bg-gray-700 text-gray-900 dark:text-white placeholder-gray-400 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 transition duration-150 ease-in-out"
                      placeholder="your.email@example.com"
                      required
                    />
                  </div>
                  
                  {/* 邮箱权限说明卡片 */}
                  <div className="mt-4 p-4 bg-yellow-50 dark:bg-yellow-900/20 rounded-lg border border-yellow-200 dark:border-yellow-800">
                    <div className="flex">
                      <div className="flex-shrink-0">
                        <AlertCircle className="h-5 w-5 text-yellow-500" />
                      </div>
                      <div className="ml-3">
                        <h3 className="text-sm font-medium text-yellow-800 dark:text-yellow-300">SMTP权限设置</h3>
                        <div className="mt-2 text-sm text-yellow-700 dark:text-yellow-400">
                          <p>真实发送邮件时，需要开启邮箱的SMTP服务并获取授权码。</p>
                          <p className="mt-1">支持的邮箱提供商：Gmail、Outlook、QQ邮箱、163邮箱、新浪邮箱等。</p>
                        </div>
                      </div>
                    </div>
                  </div>
                  
                  {/* SMTP设置（展开/收起） */}
                  <div className="mt-4">
                    <button
                      type="button"
                      onClick={() => setShowSmtpSettings(!showSmtpSettings)}
                      className="flex items-center text-sm font-medium text-blue-600 dark:text-blue-400 hover:text-blue-800 dark:hover:text-blue-300 focus:outline-none"
                    >
                      {showSmtpSettings ? (
                        <>收起 SMTP 设置 <ChevronUp className="ml-1 h-4 w-4" /></>
                      ) : (
                        <>展开 SMTP 设置 <ChevronDown className="ml-1 h-4 w-4" /></>
                      )}
                    </button>
                    
                    {showSmtpSettings && (
                      <motion.div
                        initial={{ opacity: 0, height: 0 }}
                        animate={{ opacity: 1, height: 'auto' }}
                        exit={{ opacity: 0, height: 0 }}
                        transition={{ duration: 0.3 }}
                        className="mt-4 space-y-4"
                      >
                        <div className="space-y-2">
                          <label htmlFor="smtpServer" className="block text-sm font-medium text-gray-700 dark:text-gray-300">
                            SMTP 服务器
                          </label>
                          <input
                            type="text"
                            id="smtpServer"
                            value={smtpServer}
                            onChange={(e) => setSmtpServer(e.target.value)}
                            className="block w-full px-3 py-2 border border-gray-300 dark:border-gray-600 rounded-lg bg-white dark:bg-gray-700 text-gray-900 dark:text-white placeholder-gray-400 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 transition duration-150 ease-in-out"
                            placeholder="smtp.example.com"
                          />
                        </div>
                        
                        <div className="space-y-2">
                          <label htmlFor="smtpPort" className="block text-sm font-medium text-gray-700 dark:text-gray-300">
                            SMTP 端口
                          </label>
                          <input
                            type="number"
                            id="smtpPort"
                            value={smtpPort}
                            onChange={(e) => setSmtpPort(e.target.value)}
                            className="block w-full px-3 py-2 border border-gray-300 dark:border-gray-600 rounded-lg bg-white dark:bg-gray-700 text-gray-900 dark:text-white placeholder-gray-400 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 transition duration-150 ease-in-out"
                            placeholder="587"
                          />
                        </div>
                        
                        <div className="space-y-2">
                          <label htmlFor="smtpPassword" className="block text-sm font-medium text-gray-700 dark:text-gray-300">
                            邮箱授权码（非登录密码）
                          </label>
                          <input
                            type={showPassword ? "text" : "password"}
                            id="smtpPassword"
                            value={smtpPassword}
                            onChange={(e) => setSmtpPassword(e.target.value)}
                            className="block w-full px-3 py-2 border border-gray-300 dark:border-gray-600 rounded-lg bg-white dark:bg-gray-700 text-gray-900 dark:text-white placeholder-gray-400 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 transition duration-150 ease-in-out"
                            placeholder="请输入授权码"
                          />
                          <button
                            type="button"
                            onClick={() => setShowPassword(!showPassword)}
                            className="mt-1 text-xs text-blue-600 dark:text-blue-400 hover:text-blue-800 dark:hover:text-blue-300 focus:outline-none"
                          >
                            {showPassword ? "隐藏" : "显示"} 授权码
                          </button>
                        </div>
                        
                        <div className="p-3 bg-blue-50 dark:bg-blue-900/20 rounded-lg border border-blue-100 dark:border-blue-800">
                          <p className="text-xs text-blue-700 dark:text-blue-400">
                            <strong>获取授权码方法：</strong><br />
                          1. Gmail：在账户设置中启用"两步验证"，然后生成应用专用密码<br />
                          2. QQ邮箱：设置 - 账户 - 开启SMTP服务 - 获取授权码<br />
                          3. 163邮箱：设置 - POP3/SMTP/IMAP - 开启服务 - 设置授权码<br />
                          4. 新浪邮箱：设置 - 客户端授权密码 - 开启SMTP服务 - 生成授权码
                          </p>
                        </div>
                      </motion.div>
                    )}
                  </div>
                </motion.div>

                 {/* 收件人表格上传 */}
                <motion.div variants={itemVariants} className="space-y-2">
                  <div className="flex justify-between items-center">
                    <label htmlFor="recipientsFile" className="block text-sm font-medium text-gray-700 dark:text-gray-300">
                      收件人表格 <span className="text-blue-600 dark:text-blue-400">(Excel格式)</span>
                    </label>
                    <button
                      type="button"
                      onClick={downloadExcelTemplate}
                      className="text-xs text-blue-600 dark:text-blue-400 hover:text-blue-800 dark:hover:text-blue-300 focus:outline-none flex items-center"
                    >
                      <FileSpreadsheet className="h-3 w-3 mr-1" />
                      下载Excel模板
                    </button>
                  </div>
                  <div 
                    className={`relative border-2 border-dashed rounded-lg p-8 text-center cursor-pointer transition-all duration-200 ${
                      recipientsFile 
                        ? 'border-green-500 bg-green-50 dark:bg-green-900/20' 
                        : 'border-gray-300 dark:border-gray-600 hover:border-blue-500 hover:bg-blue-50 dark:hover:bg-blue-900/20'
                    }`}
                    onClick={() => recipientsFileInputRef.current?.click()}
                  >
                    <input
                      ref={recipientsFileInputRef}
                      id="recipientsFile"
                      type="file"
                      accept=".xls,.xlsx"
                      onChange={handleRecipientsFileUpload}
                      className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
                    />
                    {recipientsFile ? (
                      <div className="flex flex-col items-center">
                        <CheckCircle className="h-12 w-12 text-green-500 mb-2" /><p className="text-sm font-medium text-gray-900 dark:text-white truncate max-w-sm">
                          {recipientsFile.name}
                        </p>
                        <p className="text-xs text-gray-500 dark:text-gray-400 mt-1">
                          {Math.round(recipientsFile.size / 1024)} KB
                        </p>
                      </div>
                    ) : (
                      <div className="flex flex-col items-center">
                        <FileSpreadsheet className="h-12 w-12 text-gray-400 mb-2" />
                        <p className="text-sm font-medium text-gray-900 dark:text-white">
                          拖放Excel文件到此处，或点击上传
                        </p>
                        <p className="text-xs text-gray-500 dark:text-gray-400 mt-1">
                          支持 .xls 和 .xlsx 格式
                        </p>
                      </div>
                    )}
                  </div>
                  <p className="text-xs text-gray-500 dark:text-gray-400 mt-1">
                    表格需包含两列：名称和邮箱
                  </p>
                </motion.div>

                {/* 邮件主题输入 */}
                <motion.div variants={itemVariants} className="space-y-2">
                  <label htmlFor="emailSubject" className="block text-sm font-medium text-gray-700 dark:text-gray-300">
                    邮件主题
                  </label>
                  <div className="relative">
                    <input
                      type="text"
                      id="emailSubject"
                      value={emailSubject}
                      onChange={(e) => setEmailSubject(e.target.value)}
                      className="block w-full pl-3 pr-3 py-3 border border-gray-300 dark:border-gray-600 rounded-lg bg-white dark:bg-gray-700 text-gray-900 dark:text-white placeholder-gray-400 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 transition duration-150 ease-in-out"
                      placeholder="请输入邮件主题"
                    />
                    {detectedSubject && (
                      <div className="mt-2 p-3 bg-green-50 dark:bg-green-900/20 rounded-lg border border-green-100 dark:border-green-800">
                        <p className="text-xs text-green-700 dark:text-green-400 flex items-center">
                          <CheckCircle className="h-4 w-4 mr-1" />
                          已从文档中检测到主题: <strong>"{detectedSubject}"</strong>
                        </p>
                      </div>
                    )}
                  </div>
                  <p className="text-xs text-gray-500 dark:text-gray-400 mt-1">
                    Word文档中红色背景的内容将自动识别为邮件主题
                  </p>
                </motion.div>

                {/* 邮件内容文档上传 */}
                <motion.div variants={itemVariants} className="space-y-2">
                  <label htmlFor="emailContentFile" className="block text-sm font-medium text-gray-700 dark:text-gray-300">
                    邮件内容文档 <span className="text-blue-600 dark:text-blue-400">(Word格式)</span>
                  </label>
                  <div 
                    className={`relative border-2 border-dashed rounded-lg p-8 text-center cursor-pointer transition-all duration-200 ${
                      emailContentFile 
                        ? 'border-green-500 bg-green-50 dark:bg-green-900/20' 
                        : 'border-gray-300 dark:border-gray-600 hover:border-blue-500 hover:bg-blue-50 dark:hover:bg-blue-900/20'
                    }`}
                    onClick={() => emailContentFileInputRef.current?.click()}
                  >
                    <input
                      ref={emailContentFileInputRef}
                      id="emailContentFile"
                      type="file"
                      accept=".doc,.docx"
                      onChange={handleEmailContentFileUpload}
                      className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
                    />
                    {emailContentFile ? (
                      <div className="flex flex-col items-center">
                        <CheckCircle className="h-12 w-12 text-green-500 mb-2" />
                        <p className="text-sm font-medium text-gray-900 dark:text-white truncate max-w-sm">
                          {emailContentFile.name}
                        </p>
                        <p className="text-xs text-gray-500 dark:text-gray-400 mt-1">
                          {Math.round(emailContentFile.size / 1024)} KB
                        </p>
                      </div>
                    ) : (
                      <div className="flex flex-col items-center">
                        <FileText className="h-12 w-12 text-gray-400 mb-2" />
                        <p className="text-sm font-medium text-gray-900 dark:text-white">
                          拖放Word文档到此处，或点击上传
                        </p>
                        <p className="text-xs text-gray-500 dark:text-gray-400 mt-1">
                          支持 .doc 和 .docx 格式
                        </p>
                      </div>
                    )}
                  </div>
                  <p className="text-xs text-gray-500 dark:text-gray-400 mt-1">
                    请用黄色高亮标记需要替换的收件人内容
                  </p>
                </motion.div>

                {/* 提交按钮 */}
                <motion.div 
                  variants={itemVariants}
                  className="flex flex-col sm:flex-row gap-4 justify-center mt-10"
                >
                  <button
                    type="submit"
                    disabled={isProcessing}
                    className={`inline-flex items-center px-8 py-3 border border-transparent text-base font-medium rounded-md shadow-sm text-white bg-blue-600 hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-blue-500 transition-all duration-200 transform hover:scale-105 disabled:opacity-50 disabled:cursor-not-allowed ${
                      isProcessing ? 'animate-pulse' : ''
                    }`}
                  >
                    {isProcessing ? (
                      <>
                        <svg className="animate-spin -ml-1 mr-2 h-5 w-5 text-white" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24">
                          <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"></circle>
                          <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path>
                        </svg>
                        处理中...
                      </>
                    ) : (
                      <>
                        开始处理 <ChevronRight className="ml-1 h-5 w-5" />
                      </>
                    )}
                  </button>
                  
                  {/* 填充测试数据按钮 */}
                  <button
                    type="button"
                    onClick={fillWithTestData}
                    className="inline-flex items-center px-6 py-3 border border-gray-300 dark:border-gray-600 text-base font-medium rounded-md shadow-sm text-gray-700 dark:text-gray-300 bg-white dark:bg-gray-700 hover:bg-gray-50 dark:hover:bg-gray-600 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-blue-500 transition-all duration-200"
                  >
                    <FileText className="mr-2 h-5 w-5" />
                    填充测试数据
                  </button>
                </motion.div>
              </motion.div>
            </form>
          </div>
        </motion.div>

        {/* 结果展示区域 */}
        {showResults && (
          <motion.div 
            id="results-section"
            className="bg-white dark:bg-gray-800 rounded-xl shadow-xl overflow-hidden border border-gray-100 dark:border-gray-700"
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.5 }}
          >
            <div className="p-8">
              <div className="flex items-center justify-between mb-6">
                <h2 className="text-2xl font-bold text-gray-900 dark:text-white">处理结果</h2>
                <button
                  onClick={resetForm}
                  className="inline-flex items-center px-4 py-2 border border-gray-300 dark:border-gray-600 rounded-md shadow-sm text-sm font-medium text-gray-700 dark:text-gray-300 bg-white dark:bg-gray-700 hover:bg-gray-50 dark:hover:bg-gray-600 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-blue-500 transition-colors duration-150"
                >
                  <RotateCcw className="h-4 w-4 mr-2" />
                  重新开始
                </button>
              </div>
              
              <div className="overflow-x-auto">
                <table className="min-w-full divide-y divide-gray-200 dark:divide-gray-700">
                  <thead className="bg-gray-50 dark:bg-gray-700">
                    <tr>
                        <th scope="col" className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase tracking-wider">
                          收件人名称
                        </th>
                        <th scope="col" className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase tracking-wider">
                          收件人邮箱
                        </th>
                        <th scope="col" className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase tracking-wider">
                          邮件主题
                        </th>
                        <th scope="col" className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase tracking-wider">
                          状态
                        </th>
                      <th scope="col" className="px-6 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-300 uppercase tracking-wider">
                        消息
                      </th>
                    </tr>
                  </thead>
                  <tbody className="bg-white dark:bg-gray-800 divide-y divide-gray-200 dark:divide-gray-700">
                    {results.map((result, index) => (
                      <motion.tr 
                        key={index}
                        initial={{ opacity: 0, y: 10 }}
                        animate={{ opacity: 1, y: 0 }}
                        transition={{ duration: 0.3, delay: index * 0.1 }}
                        className={cn(
                          "hover:bg-gray-50 dark:hover:bg-gray-700 transition-colors duration-150",
                          index % 2 === 0 ? "bg-white dark:bg-gray-800" : "bg-gray-50 dark:bg-gray-700/50"
                        )}
                      >
                        <td className="px-6 py-4 whitespace-nowrap text-sm font-medium text-gray-900 dark:text-white">
                          {result.recipient.name}
                        </td>
                        <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-500 dark:text-gray-300">
                          {result.recipient.email}
                        </td>
                        <td className="px-6 py-4 text-sm text-gray-500 dark:text-gray-300 max-w-xs truncate" title={result.subject}>
                          {result.subject || '无主题'}
                        </td>
                        <td className="px-6 py-4 whitespace-nowrap">
                          <span className={`px-2 inline-flex text-xs leading-5 font-semibold rounded-full ${
                            result.success 
                              ? 'bg-green-100 text-green-800 dark:bg-green-900/30 dark:text-green-400' 
                              : 'bg-red-100 text-red-800 dark:bg-red-900/30 dark:text-red-400'
                          }`}>
                            {result.success ? '成功' : '失败'}
                          </span>
                        </td>
                        <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-500 dark:text-gray-300">
                          {result.message}
                        </td>
                      </motion.tr>
                    ))}
                  </tbody>
                </table>
              </div>
              
              <div className="mt-8 p-4 bg-blue-50 dark:bg-blue-900/20 rounded-lg border border-blue-100 dark:border-blue-800">
                <div className="flex">
                  <div className="flex-shrink-0">
                    <InfoIcon className="h-5 w-5 text-blue-500" />
                  </div>
                  <div className="ml-3">
                    <h3 className="text-sm font-medium text-blue-800 dark:text-blue-300">处理统计</h3>
               <div className="mt-2 text-sm text-blue-700 dark:text-blue-400">
                        <p>总收件人数：{results.length}</p>
                        <p>成功发送：{results.filter(r => r.success).length}</p>
                        <p>发送失败：{results.filter(r => !r.success).length}</p>
                        <p className="mt-1 text-xs text-blue-600 dark:text-blue-500">
                          邮件主题：{results[0]?.subject || '无主题'}
                        </p>
                       <div className="mt-3 p-3 bg-yellow-50 dark:bg-yellow-900/20 rounded-lg border border-yellow-200 dark:border-yellow-800">
                         <p className="text-xs text-yellow-700 dark:text-yellow-400">
                           <strong>演示模式说明：</strong><br />
                           当前为演示版本，显示的是基于您上传文件生成的模拟数据。<br />
                           实际应用中，系统将使用xlsx库解析您的Excel文件，提取真实的收件人信息。<br />
                           如需完整功能，请集成xlsx库并实现真实的文件解析逻辑。
                         </p>
                       </div>
                     </div>
                  </div>
                </div>
              </div>
            </div>
          </motion.div>
        )}
      </div>
    </div>
  );
}

// 下载Excel模板函数
import * as XLSX from 'xlsx';

function downloadExcelTemplate() {
  try {
    // 创建工作簿
    const wb = XLSX.utils.book_new();
    
    // 创建表头和示例数据
    const data = [
      ['名称', '邮箱'],
      ['张三', 'zhangsan@example.com'],
      ['李四', 'lisi@example.com'],
      ['王五', 'wangwu@example.com']
    ];
    
    // 创建工作表
    const ws = XLSX.utils.aoa_to_sheet(data);
    
    // 设置列宽
    const colWidths = [
      { wch: 15 }, // 名称列
      { wch: 30 }  // 邮箱列
    ];
    ws['!cols'] = colWidths;
    
    // 添加工作表到工作簿
    XLSX.utils.book_append_sheet(wb, ws, '收件人列表');
    
    // 生成Excel文件并下载
    XLSX.writeFile(wb, '收件人模板.xlsx');
    
    // 显示成功提示
    toast.success('Excel模板已下载');
  } catch (error) {
    console.error('下载模板失败:', error);
    // 降级为CSV格式
    try {
      // CSV内容
      const csvContent = [
        '名称,邮箱',
        '张三,zhangsan@example.com',
        '李四,lisi@example.com',
        '王五,wangwu@example.com'
      ].join('\n');
      
      // 创建Blob对象
      const blob = new Blob(['\ufeff' + csvContent], { type: 'text/csv;charset=utf-8;' });
      
      // 创建下载链接
      const link = document.createElement('a');
      const url = URL.createObjectURL(blob);
      
      link.setAttribute('href', url);
      link.setAttribute('download', '收件人模板.csv');
      link.style.visibility = 'hidden';
      
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      
      // 显示成功提示
      toast.success('CSV模板已下载（可用Excel打开）');
    } catch (csvError) {
      console.error('下载CSV模板失败:', csvError);
      toast.error('下载模板失败，请重试');
    }
  }
}