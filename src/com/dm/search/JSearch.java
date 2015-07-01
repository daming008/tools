package com.dm.search;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileFilter;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.UnsupportedEncodingException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class JSearch {
	private static String path = "C:\\wamp\\www\\upload2";
	private static String filename = ".*\\.php";
	private static String targetStr = "Pw_verify_code";
	private static String charset = "";
	private static boolean isSearchSubdirectory = true;
	private static String usage = "Usage: java JSearch -t str [-options]\noptions:\n\t -d directory\t 要查找的子目录，默认是当前目录\n\t -f file\t 要查找的文件，默认所有txt和sql文件，支持正则\n\t -t targetStr\t 将要查找的目标字符串，不可空，支持正则\n\t -c charset\t 字符编码\n\t -r\t\t 递归查找子目录\n示例：\njava JSearch -t t00ls\njava JSearch -t t00ls -d /root/test/txt -f .*\\.txt -c gbk -r";
	
	public static void main(String[] args) {
//		if(!parseArgs(args)){
//			return;
//		}
		List<File> list = getFiles(path, filename, isSearchSubdirectory);
		find(list);
	}
	
	/**
	 * 在文件中查找字符串并输出
	 * @param list 文件列表
	 */
	private static void find(List<File> list) {
		for (File file : list) {
			String result = find(file.getAbsolutePath());
			if("".equals(result)){
				continue;
			}
			System.out.println("**********在" + file.getAbsolutePath() + "中找到：**********");
			System.out.println(result);
		}
	}
	
	/**
	 * 在文件中查找字符串
	 * @param filename 文件路径
	 * @return 查找结果
	 */
	private static String find(String filename) {
		if(targetStr == null || targetStr.length() == 0) {
			throw new RuntimeException("目标字符串不能为空");
		}
		BufferedReader br = null;
		InputStreamReader isr = null;
		FileInputStream fis = null;
		String line = null;
		String result = "";
		try {
			fis = new FileInputStream(filename);
			isr = "".equals(charset) ? new InputStreamReader(fis) : new InputStreamReader(fis, charset);
			br = new BufferedReader(isr);
			
			Pattern pattern = Pattern.compile(targetStr);
			
			while((line = br.readLine()) != null){
				Matcher matcher = pattern.matcher(line);
				if(matcher.find()){
					result += line + "\n";
				}
			}
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (UnsupportedEncodingException e) {
			System.out.println("编码无效");
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			if(fis != null){
				try {
					fis.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
			if(isr != null){
				try {
					isr.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
			if(br != null){
				try {
					br.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return result;
	}

	/**
	 * 获取某一路径下的所有符合条件的文件
	 * @param path 路径
	 * @param filename 文件名过滤条件
	 * @param isSearchSubdirectory 是否查找子目录
	 * @return 文件列表
	 */
	private static List<File> getFiles(String path, final String filename, boolean isSearchSubdirectory) {
		List<File> result = new ArrayList<File>();
		File file = new File(path).getAbsoluteFile();
		if(!file.exists()){
			return result;
		}
		if(file.isFile()){
			result.add(file);
			return result;
		}
		
		File[] files = file.listFiles(new FileFilter() {
			
			@Override
			public boolean accept(File file) {
				if(!file.getAbsoluteFile().isDirectory()){
					if("".equals(filename)){
						return true;
					}
					if(file.getAbsoluteFile().getName().matches(filename)){
						return true;
					}
				}
				return false;
			}
			
		});
		result.addAll(Arrays.asList(files));
		
		if(isSearchSubdirectory){
			File[] directorys = file.listFiles(new FileFilter() {
				
				@Override
				public boolean accept(File file) {
					if(file.getAbsoluteFile().isDirectory()){
						return true;
					}
					return false;
				}
				
			});
			
			for (File directory : directorys) {
				List<File> list = getFiles(directory.getAbsolutePath(), filename, isSearchSubdirectory);
				result.addAll(list);
			}
		}
		
		return result;
	}

	/**
	 * 解析参数
	 * @param args 参数数组
	 * @return 解析是否成功
	 
	private static boolean parseArgs(String[] args) {
		if(args.length == 0){
			System.out.println(usage);
			return false;
		}
		try {
			for (int i = 0; i < args.length; i++) {
				if("-d".equals(args[i])){
					path = args[i + 1];
				}
				if("-f".equals(args[i])){
					filename = args[i + 1];
				}
				if("-t".equals(args[i])){
					targetStr = args[i + 1];
				}
				if("-c".equals(args[i])){
					charset = args[i + 1];
				}
				if("-r".equals(args[i])){
					isSearchSubdirectory = true;
				}
			}
			if("".equals(targetStr)){
				System.out.println("请用-t参数指定要查找的字符串");
				return false;
			}
			return true;
		} catch (Exception e) {
			System.out.println(usage);
			return false;
		}
	}
	*/
}