import asyncio
import random
import aiohttp

class RequestManager:
    def __init__(self, user_agents=None):
        if user_agents is None:
            self.user_agents = [
                'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/111.0.5563.65 Safari/537.36',
                'Mozilla/5.0 (iPhone; CPU iPhone OS 16_2 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/16.2 Mobile/15E148 Safari/604.1',
                'Mozilla/5.0 (Linux; Android 13; SM-G960F) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/111.0.5563.65 Mobile Safari/537.36',
                'Mozilla/5.0 (Macintosh; Intel Mac OS X 12_6) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/16.2 Safari/605.1.15',
                'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:110.0) Gecko/20100101 Firefox/110.0'
            ]
        else:
            self.user_agents = user_agents

    async def make_request(self, url, timeout=10):
        user_agent = random.choice(self.user_agents)
        headers = {'User-Agent': user_agent}
        try:
            async with aiohttp.ClientSession() as session:
                async with session.get(url, headers=headers, timeout=timeout) as response:
                    status = response.status
                    text = await response.text()
                    return status, text
        except aiohttp.ClientError as e:
            return None, str(e)
        except asyncio.TimeoutError:
            return None, "Timeout error"

if __name__ == "__main__":
    async def main():
        request_manager = RequestManager()
        status, text = await request_manager.make_request('https://httpbin.org/headers')
        print(f"Status: {status}")
        print(f"Text: {text}")

    asyncio.run(main())
