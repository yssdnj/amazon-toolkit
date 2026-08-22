import unittest
from unittest.mock import Mock, patch

from tools import asin_variant_score


class XiyouV2ApiTests(unittest.TestCase):
    @patch.object(asin_variant_score, "XIYOU_API_KEY", "test-v2-api-key")
    @patch("tools.asin_variant_score.requests.post")
    def test_api_call_uses_v2_base_url_and_api_key_headers(self, post):
        response = Mock()
        response.raise_for_status.return_value = None
        response.json.return_value = {"entities": []}
        post.return_value = response

        body = {"country": "US", "asin": "B0B2RM68G2"}
        result = asin_variant_score._api_call("/v1/asins/variations", body)

        self.assertEqual({"entities": []}, result)
        post.assert_called_once_with(
            "https://openapi.xydc.com/v1/asins/variations",
            headers={
                "Content-Type": "application/json",
                "X-Auth-Version": "2.0",
                "X-Api-Key": "test-v2-api-key",
            },
            json=body,
            timeout=30,
        )


if __name__ == "__main__":
    unittest.main()
