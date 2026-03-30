import queue
import unittest
from unittest.mock import patch

import main


class FakeResponse:
	def __init__(self, payload):
		self.payload = payload

	def json(self):
		return self.payload


class NewEndpointFetchTests(unittest.TestCase):
	def test_fetch_robot_errors_data_flattens_records(self):
		responses = [
			FakeResponse(
				{
					"results": [
						{
							"date": "2026-03-24",
							"version": "1.3.0",
							"result": {
								"robot_errors": [
									{"robot_id": 5, "robot_error_code": 218, "local_installation_timestamp": "2026-03-24 12:16:16"},
									{"robot_id": 5, "robot_error_code": 219, "local_installation_timestamp": "2026-03-24 12:16:25"},
								]
							},
						},
					],
					"next": "page-2",
				}
			),
			FakeResponse(
				{
					"results": [
						{
							"date": "2025-12-31",
							"version": "1.3.0",
							"result": {"robot_errors": [{"robot_id": 99}]},
						},
					],
					"next": None,
				}
			),
		]

		with patch("main.api_get", side_effect=responses):
			df = main.fetch_robot_errors_data("installation", 2026, queue.Queue())

		self.assertEqual(len(df), 2)
		self.assertEqual(df["robot_error_code"].tolist(), [218, 219])
		self.assertTrue((df["date"] == "2026-03-24").all())

	def test_fetch_robot_mtbf_data_expands_active_times(self):
		responses = [
			FakeResponse(
				{
					"results": [
						{
							"date": "2026-03-24",
							"version": "1.4.10",
							"result": {
								"robot_mtbf": {
									"robot_mtbf": 62364.9,
									"total_errors": 20,
									"total_time_active_s": 1247298,
									"robot_mtbf_system_stops": 623649.0,
									"total_errors_system_stops": 2,
								},
								"robot_active_times": [
									{"robot_id": 19, "total_time_active_s": 39829},
									{"robot_id": 27, "total_time_active_s": 40392},
								],
							},
						},
					],
					"next": None,
				}
			),
		]

		with patch("main.api_get", side_effect=responses):
			df = main.fetch_robot_mtbf_data("installation", 2026, queue.Queue())

		self.assertEqual(len(df), 2)
		self.assertEqual(df["robot_id"].tolist(), [19, 27])
		self.assertTrue((df["total_errors"] == 20).all())

	def test_fetch_incidents_data_joins_details(self):
		responses = [
			FakeResponse(
				{
					"results": [
						{
							"date": "2026-03-24",
							"version": "1.1.7",
							"result": {
								"incidents": [
									{
										"incident_id": 7606,
										"status": "RESOLVED",
										"details_display_name": ["A", "B"],
										"start_local_timestamp": "2026-03-24 12:16:52",
									}
								]
							},
						},
					],
					"next": None,
				}
			),
		]

		with patch("main.api_get", side_effect=responses):
			df = main.fetch_incidents_data("installation", 2026, queue.Queue())

		self.assertEqual(len(df), 1)
		self.assertEqual(df.loc[0, "incident_id"], 7606)
		self.assertEqual(df.loc[0, "details_display_name"], "A | B")

	def test_fetch_uptime_trend_data_returns_one_row_per_record(self):
		responses = [
			FakeResponse(
				{
					"results": [
						{
							"date": "2026-03-22",
							"version": "1.0.1",
							"result": {
								"uptime_trend": {
									"week": 12,
									"year": 2026,
									"trend": 0.99825,
									"outlier": True,
									"lower_bound": 0.99535,
									"upper_bound": 1.0,
									"weekly_uptime": 0.99434,
									"long_term_trend": "STABLE",
									"short_term_trend": "DOWN",
								}
							},
						},
					],
					"next": None,
				}
			),
		]

		with patch("main.api_get", side_effect=responses):
			df = main.fetch_uptime_trend_data("installation", 2026, queue.Queue())

		self.assertEqual(len(df), 1)
		self.assertEqual(df.loc[0, "week"], 12)
		self.assertEqual(df.loc[0, "short_term_trend"], "DOWN")


if __name__ == "__main__":
	unittest.main()