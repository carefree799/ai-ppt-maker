import os
import re
import json
import uuid
import requests
from flask import Flask, render_template, request, session, redirect, url_for

app = Flask(__name__)
app.config['OUTPUT_FOLDER'] = os.path.join(os.path.dirname(__file__), 'output')
app.config['SECRET_KEY'] = os.urandom(24)
app.config['SESSION_DATA_DIR'] = '/tmp/ai-ppt-maker'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)
os.makedirs(app.config['SESSION_DATA_DIR'], exist_ok=True)


def load_api_key():
    settings_path = os.path.join(os.path.dirname(__file__), 'settings.json')
    with open(settings_path, 'r') as f:
        settings = json.load(f)
    return settings.get('ANTHROPIC_API_KEY', '')


def generate_outline(topic, description=''):
    api_key = load_api_key()
    if not api_key:
        raise ValueError('API Key 未配置，请在 settings.json 中设置 ANTHROPIC_API_KEY')

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

    response = requests.post(
        'https://api.minimaxi.com/anthropic/v1/messages',
        headers={
            'Authorization': f'Bearer {api_key}',
            'Content-Type': 'application/json'
        },
        json={
            'model': 'MiniMax-M2.1',
            'max_tokens': 2000,
            'messages': [{'role': 'user', 'content': prompt}]
        },
        timeout=60
    )

    if response.status_code != 200:
        raise ValueError(f'API 请求失败: {response.status_code}')

    result = response.json()
    content = result.get('content', [{}])[0].get('text', '')

    json_match = re.search(r'\{[\s\S]*\}', content)
    if json_match:
        outline = json.loads(json_match.group())

        if 'slides' in outline and isinstance(outline['slides'], list):
            for i, slide in enumerate(outline['slides']):
                slide['page'] = i + 1

        return outline
    raise ValueError(f'无法解析 AI 输出: {content[:200]}...' if len(content) > 200 else f'无法解析 AI 输出: {content}')


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


if __name__ == '__main__':
    app.run(debug=True)