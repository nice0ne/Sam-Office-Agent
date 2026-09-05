import { OpenAICompatibleProvider } from './openai';

export class GLMProvider extends OpenAICompatibleProvider {
  constructor(baseUrl: string = 'https://open.bigmodel.cn/api/paas/v4') {
    super('glm', 'GLM Coding Global', baseUrl);
  }
}
