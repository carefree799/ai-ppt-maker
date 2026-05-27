import os
import re
import json
import uuid
import requests
from flask import Flask, render_template, request, session, redirect, url_for, send_from_directory

app = Flask(__name__)
app.config['OUTPUT_FOLDER'] = os.path.join(os.path.dirname(__file__), 'static')
app.config['SECRET_KEY'] = os.urandom(24)
app.config['SESSION_DATA_DIR'] = '/tmp/ai-ppt-maker'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)
os.makedirs(app.config['SESSION_DATA_DIR'], exist_ok=True)


def load_api_key():
    # 先从环境变量读取，没有则从 settings.json 读取
    api_key = os.environ.get('ANTHROPIC_API_KEY', '')
    if api_key:
        return api_key

    settings_path = os.path.join(os.path.dirname(__file__), 'settings.json')
    with open(settings_path, 'r') as f:
        settings = json.load(f)
    return settings.get('ANTHROPIC_API_KEY', '')


def call_minimax_api(prompt, max_tokens=2000):
    """调用 MiniMax API"""
    api_key = load_api_key()
    if not api_key:
        raise ValueError('API Key 未配置，请在 settings.json 中设置 ANTHROPIC_API_KEY')

    try:
        response = requests.post(
            'https://api.minimaxi.com/anthropic/v1/messages',
            headers={
                'X-Api-Key': api_key,
                'Content-Type': 'application/json'
            },
            json={
                'model': 'MiniMax-M2.1',
                'max_tokens': max_tokens,
                'messages': [{'role': 'user', 'content': prompt}]
            },
            timeout=60
        )
    except requests.exceptions.Timeout:
        raise ValueError('API 请求超时，请重试')
    except requests.exceptions.RequestException as e:
        raise ValueError(f'API 请求失败: {str(e)}')

    if response.status_code != 200:
        try:
            error_msg = response.json().get('error', {}).get('message', '') or response.text
        except:
            error_msg = response.text
        raise ValueError(f'API 错误 ({response.status_code}): {error_msg[:100]}')

    try:
        result = response.json()
    except json.JSONDecodeError:
        raise ValueError('API 返回非 JSON 格式')

    content = result.get('content', [])
    for item in content:
        if item.get('type') == 'text':
            return item.get('text', '')
    return ''


def generate_outline(topic, description=''):
    prompt = f"""你是一个 PPT 结构助手。根据用户的主题生成一个 PPT 大纲。

主题：{topic}
补充说明：{description}

请以 JSON 格式返回大纲，必须包含以下字段：
- title: PPT 标题
- slides: 数组，每页包含 page(页码，从1开始) 和 title(标题)

要求：
- 5-10 页为宜
- 逻辑清晰，符合主题
- 用中文回复
- 只返回 JSON，不要任何解释"""

    content = call_minimax_api(prompt, max_tokens=2000)

    json_match = re.search(r'\{[\s\S]*\}', content)
    if json_match:
        try:
            outline = json.loads(json_match.group())

            if 'slides' in outline and isinstance(outline['slides'], list):
                for i, slide in enumerate(outline['slides']):
                    slide['page'] = i + 1

            return outline
        except json.JSONDecodeError:
            raise ValueError(f'无法解析 AI 输出的 JSON')
    raise ValueError(f'无法解析 AI 输出: {content[:200]}...' if len(content) > 200 else f'无法解析 AI 输出: {content}')


def generate_slides_content(session_data):
    topic = session_data['topic']
    description = session_data.get('description', '')
    outline = session_data['outline']

    prompt = f"""你是一个 PPT 内容助手。根据大纲为每一页生成详细内容。

主题：{topic}
补充说明：{description}

大纲：
{json.dumps(outline, ensure_ascii=False)}

请为每一页生成以下内容（返回 JSON 数组，必须包含以下字段）：
- title: 标题
- bullets: 要点数组（3-5个）
- layout: 布局类型，可选值：title_slide(封面)、single_column(单栏)、two_column(双栏对比)、comparison(左右对比)、chart(图表页)
- theme: 主题色彩，可选值：blue(蓝色)、green(绿色)、purple(紫色)、gray(灰色)，第一页用 blue，其他随机或与主题匹配
- chart_type: 如果 layout 是 chart，可选 bar、pie、line
- chart_data: 如果有图表，格式如 {{"labels":["A","B"],"values":[10,20]}}
- notes: 演讲备注（可选）

用中文回复，只返回 JSON 格式，不要任何解释"""

    content = call_minimax_api(prompt, max_tokens=4000)

    json_match = re.search(r'\[[\s\S]*\]', content)
    if json_match:
        try:
            return json.loads(json_match.group())
        except json.JSONDecodeError:
            raise ValueError(f'无法解析 AI 输出的 JSON')
    raise ValueError(f'无法解析 AI 输出: {content[:200]}...' if len(content) > 200 else f'无法解析 AI 输出: {content}')


def create_pptx(slides_data, output_path):
    """使用专业级渲染器创建 PPT"""
    from ppt_renderer import render_pptx

    if not slides_data:
        raise ValueError('没有幻灯片内容')

    return render_pptx(slides_data, output_path)


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/generate', methods=['POST'])
def generate():
    topic = request.form.get('topic', '').strip()
    description = request.form.get('description', '').strip()

    if not topic:
        return render_template('index.html', error='请输入主题')

    try:
        outline = generate_outline(topic, description)
        session_id = str(uuid.uuid4())

        session_data = {
            'topic': topic,
            'description': description,
            'outline': outline
        }
        session_file = os.path.join(app.config['SESSION_DATA_DIR'], f'{session_id}.json')
        with open(session_file, 'w') as f:
            json.dump(session_data, f, ensure_ascii=False)

        return render_template('preview.html', outline=outline, topic=topic,
                         description=description, session_id=session_id)
    except Exception as e:
        return render_template('index.html', error=f'生成失败: {str(e)}', topic=topic, description=description)


@app.route('/create', methods=['POST'])
def create():
    session_id = request.form.get('session_id', '')

    if not session_id:
        return render_template('index.html', error='无效的会话')

    session_file = os.path.join(app.config['SESSION_DATA_DIR'], f'{session_id}.json')
    if not os.path.exists(session_file):
        return render_template('index.html', error='会话已过期，请重新输入')

    with open(session_file, 'r') as f:
        session_data = json.load(f)

    try:
        slides_content = generate_slides_content(session_data)

        if not isinstance(slides_content, list):
            raise ValueError('生成的幻灯片内容格式不正确')

        output_file = f'{session_id}.pptx'
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_file)
        create_pptx(slides_content, output_path)

        return render_template('download.html', filename=output_file)
    except Exception as e:
        return render_template('index.html', error=f'生成失败: {str(e)}')


@app.route('/create-pptx', methods=['POST'])
def generate_pptx():
    session_id = request.form.get('session_id', '')

    if not session_id:
        return render_template('index.html', error='无效的会话')

    session_file = os.path.join(app.config['SESSION_DATA_DIR'], f'{session_id}.json')
    if not os.path.exists(session_file):
        return render_template('index.html', error='会话已过期，请重新输入')

    with open(session_file, 'r') as f:
        session_data = json.load(f)

    try:
        slides_content = generate_slides_content(session_data)

        if not isinstance(slides_content, list):
            raise ValueError('生成的幻灯片内容格式不正确')

        output_file = f'{session_id}.pptx'
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_file)
        create_pptx(slides_content, output_path)

        return render_template('download.html', filename=output_file)
    except Exception as e:
        return render_template('index.html', error=f'生成失败: {str(e)}')


@app.route('/download/<path:filename>')
def download(filename):
    return send_from_directory(app.config['OUTPUT_FOLDER'], filename, as_attachment=True)


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)