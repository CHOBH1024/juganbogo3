let _pipe: any = null;

export function isBrowserAIReady(): boolean {
  return !!_pipe;
}

export async function initBrowserAI(onProgress?: (pct: number) => void): Promise<void> {
  if (_pipe) return;

  // 동적 import — 사용할 때만 ONNX 런타임 로드 (메인 번들 영향 없음)
  const { pipeline, env } = await import('@huggingface/transformers');
  env.allowLocalModels = false;
  env.useBrowserCache = true;

  _pipe = await (pipeline as any)('text-generation', 'onnx-community/Qwen2.5-0.5B-Instruct', {
    dtype: 'q4',
    progress_callback: (info: any) => {
      if (info.status === 'progress' && typeof info.progress === 'number') {
        onProgress?.(Math.round(info.progress));
      }
    },
  });
}

export async function runBrowserAI(prompt: string): Promise<string> {
  if (!_pipe) throw new Error('AI 모델이 준비되지 않았습니다. 먼저 다운로드하세요.');
  const messages = [{ role: 'user' as const, content: prompt }];
  const result: any = await _pipe(messages, { max_new_tokens: 1024, do_sample: false });
  const out = result?.[0]?.generated_text;
  if (Array.isArray(out)) return out.at(-1)?.content ?? '';
  return String(out ?? '');
}
