import { Command } from 'commander';
import * as fs from 'fs';
import * as path from 'path';
import matter from 'gray-matter';
import * as XLSX from 'xlsx';

// 创建日志目录
const LOG_DIR = './logs';
if (!fs.existsSync(LOG_DIR)) {
  fs.mkdirSync(LOG_DIR, { recursive: true });
}

// 获取格式化的时间戳
function getTimestamp(): string {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');
  const hours = String(now.getHours()).padStart(2, '0');
  const minutes = String(now.getMinutes()).padStart(2, '0');
  return `${year}-${month}-${day} ${hours}-${minutes}`;
}

// 创建日志写入流
const timestamp = getTimestamp();
const logStream = fs.createWriteStream(path.join(LOG_DIR, `conversion-${timestamp}.log`));

// 日志写入函数
function writeLog(message: string, type: 'info' | 'error' = 'info'): void {
  const prefix = type === 'error' ? '❌ ' : '✨ ';
  const logMessage = `[${new Date().toISOString()}] ${prefix}${message}\n`;
  logStream.write(logMessage);
  if (type === 'error') {
    console.error(message);
  } else {
    console.log(message);
  }
}

interface Config {
  sourceDir: string;
  outputDir: string;
  ignoreDirs: string[];
}


const program = new Command();

program
  .option('-s, --source <dir>', '指定MDX文件目录')
  .option('-o, --output <dir>', '指定输出Excel文件目录')
  .option('-i, --ignore <dirs...>', '指定要忽略的目录', [])
  .parse(process.argv);

const options = program.opts();

const config: Config = {
  sourceDir: options.source || './content',
  outputDir: options.output || './excel',
  ignoreDirs: options.ignore || []
};

// 确保输出目录存在
if (!fs.existsSync(config.outputDir)) {
  fs.mkdirSync(config.outputDir, { recursive: true });
}

// 递归获取所有MDX文件
function getMdxFiles(dir: string): string[] {
  const files: string[] = [];
  const items = fs.readdirSync(dir);

  for (const item of items) {
    const fullPath = path.join(dir, item);
    const isIgnored = config.ignoreDirs.some(ignoreDir =>
      fullPath.includes(path.normalize(ignoreDir))
    );

    if (isIgnored) continue;

    if (fs.statSync(fullPath).isDirectory()) {
      files.push(...getMdxFiles(fullPath));
    } else if (item.endsWith('.mdx') || item.endsWith('.md')) {
      files.push(fullPath);
    }
  }

  return files;
}

// 处理单个MDX文件
// Excel单元格字符限制
const EXCEL_CELL_LIMIT = 32767;

function splitLongContent(content: string): string[] {
  const chunks: string[] = [];
  let remainingContent = content;

  while (remainingContent.length > 0) {
    if (remainingContent.length <= EXCEL_CELL_LIMIT) {
      chunks.push(remainingContent);
      break;
    }

    // 在EXCEL_CELL_LIMIT位置之前找到最后一个段落分隔符
    let splitIndex = remainingContent.lastIndexOf('\n\n', EXCEL_CELL_LIMIT);
    if (splitIndex === -1 || splitIndex === 0) {
      // 如果找不到段落分隔符，则在EXCEL_CELL_LIMIT位置之前找到最后一个换行符
      splitIndex = remainingContent.lastIndexOf('\n', EXCEL_CELL_LIMIT);
    }
    if (splitIndex === -1 || splitIndex === 0) {
      // 如果仍然找不到合适的分割点，就在EXCEL_CELL_LIMIT处直接截断
      splitIndex = EXCEL_CELL_LIMIT;
    }

    chunks.push(remainingContent.slice(0, splitIndex).trim());
    remainingContent = remainingContent.slice(splitIndex).trim();
  }

  return chunks;
}

interface SimpleProcessData {
  title: string;
  content: string;
}

function processLongContent(title: string, content: string): SimpleProcessData[] {
  if (content.length <= EXCEL_CELL_LIMIT) {
    return [{ title, content }];
  }

  const chunks = splitLongContent(content);
  return chunks.map((chunk, index) => ({
    title: index === 0 ? title : `${title}-${index + 1}`,
    content: chunk
  }));
}

function processMdxFile(filePath: string): SimpleProcessData[] {
  const fileContent = fs.readFileSync(filePath, 'utf-8');
  const { data, content } = matter(fileContent);
  
  const results: SimpleProcessData[] = [];
  
  // 尝试从内容中提取主标题
  const h1Match = content.match(/^#\s+(.+)$/m);
  const mainTitle = h1Match ? h1Match[1] : (data.title || path.basename(filePath, '.mdx'));
  
  // 获取所有二级标题和对应内容
  const sections = content.split(/(?=^##\s+.+$)/m).filter(Boolean) || [];
  
  // 处理主要内容（不包含二级标题的部分）
  if (sections.length > 0 && !sections[0].startsWith('##')) {
    const mainContent = sections[0].trim();
    if (mainContent) {
      results.push(...processLongContent(mainTitle, mainContent));
    }
    sections.shift(); // 移除已处理的主要内容
  }
  
  // 处理每个二级标题部分
  sections.forEach((section) => {
    const h2TitleMatch = section.match(/^##\s+(.+)$/m);
    if (h2TitleMatch) {
      const sectionTitle = h2TitleMatch[1];
      const sectionContent = section.replace(/^##\s+.+$/m, '').trim();
      
      if (sectionContent) {
        results.push(...processLongContent(`${mainTitle}-${sectionTitle}`, sectionContent));
      }
    }
  });
  
  return results;
}

// 转换MDX文件到Excel
function convertToExcel(mdxFiles: string[]): void {
  const successfulData: SimpleProcessData[] = [];
  
  for (const file of mdxFiles) {
    try {
      writeLog(`正在处理文件: ${file}`);
      const processedDataArray = processMdxFile(file);
      successfulData.push(...processedDataArray);
      writeLog(`文件处理成功: ${file}`);
    } catch (error) {
      writeLog(`文件处理失败: ${file}`, 'error');
      writeLog(`错误信息: ${error.message}`, 'error');
      continue;
    }
  }
  
  if (successfulData.length === 0) {
    writeLog('没有成功处理任何文件，无法生成Excel', 'error');
    return;
  }
  
  const worksheet = XLSX.utils.json_to_sheet(successfulData);
  const workbook = XLSX.utils.book_new();
  
  XLSX.utils.book_append_sheet(workbook, worksheet, 'MDX Content');
  
  const outputPath = path.join(config.outputDir, `output-${timestamp}.xlsx`);
  XLSX.writeFile(workbook, outputPath);
  
  writeLog(`转换完成！成功处理 ${successfulData.length}/${mdxFiles.length} 个文件`);
  writeLog(`Excel文件已保存到: ${outputPath}`);
  
  // 关闭日志流
  logStream.end();
}

try {
  const mdxFiles = getMdxFiles(config.sourceDir);
  if (mdxFiles.length === 0) {
    console.log('未找到MDX文件');
    process.exit(0);
  }
  
  convertToExcel(mdxFiles);
} catch (error) {
  console.error('转换过程中发生错误:', error);
  process.exit(1);
}