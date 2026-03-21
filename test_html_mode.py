import requests
import json
import time

# 测试文件
test_file = '/Users/dalabengba/Desktop/试卷/2025年高考全国一卷数学真题.docx'

print('=== 测试 HTML 模式完整流程 ===\n')

# 1. 上传文件
print('1. 上传文件...')
with open(test_file, 'rb') as f:
    files = {'file': f}
    data = {'mode': 'html'}
    
    response = requests.post('http://localhost:3000/api/v1/upload', files=files, data=data, timeout=10)
    result = response.json()
    
    print(f'上传响应: {json.dumps(result, indent=2, ensure_ascii=False)}')
    
    if result.get('success'):
        file_id = result['file_id']
        
        # 2. 轮询进度
        print(f'\n2. 轮询解析进度 (file_id: {file_id})...')
        for i in range(60):  # 最多等待 60 秒
            progress_response = requests.get(f'http://localhost:3000/api/v1/upload/progress/{file_id}')
            progress_data = progress_response.json()
            
            print(f'  [{i+1}s] 状态: {progress_data.get("status")}, 进度: {progress_data.get("progress")}%, 消息: {progress_data.get("message")}')
            
            if progress_data.get('status') == 'completed':
                print(f'\n✅ 解析完成!')
                print(f'模式: {progress_data.get("mode")}')
                
                if progress_data.get('mode') == 'html':
                    html_length = len(progress_data.get('html', ''))
                    print(f'HTML 长度: {html_length} 字符')
                    
                    # 保存 HTML
                    if html_length > 0:
                        with open('/Users/dalabengba/Desktop/math-demo/test_frontend.html', 'w', encoding='utf-8') as f:
                            f.write(progress_data['html'])
                        print(f'HTML 已保存到: test_frontend.html')
                    else:
                        print('⚠️ HTML 内容为空!')
                else:
                    print(f'题目数量: {len(progress_data.get("questions", []))}')
                break
            elif progress_data.get('status') == 'error':
                print(f'\n❌ 解析失败: {progress_data.get("message")}')
                break
            
            time.sleep(1)
    else:
        print(f'❌ 上传失败: {result.get("message")}')
