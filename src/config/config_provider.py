import os
import json
import shutil


class ConfigProvider:
    """
    Class for loading configuration from a per-user JSON file in AppData,
    or a custom path if provided.
    """

    @staticmethod
    def load_config_json(path=None):
        # TO_BE_CHANGED
        path = os.path.join(os.getcwd(), "config.json")
        if not os.path.exists(path):
            path = os.path.join(os.path.dirname(os.getcwd()), "config.json")

        try:
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                print(data)  # For debugging
                return data
        except Exception as e:
            print(f"Error reading config.json: {e}")
            return {}

        # """
        # Load the configuration from the user's JSON file or a specified path.
        #
        # :param path: Optional full path to config.json
        # :return: The loaded configuration as a dictionary.
        # """
        # if path is None:
        #     # --- Determine per-user AppData path ---
        #     appdata = os.getenv('APPDATA') or os.path.expanduser('~\\AppData\\Roaming')
        #     user_config_folder = os.path.join(appdata, "TO_BE_CHANGED")
        #     os.makedirs(user_config_folder, exist_ok=True)
        #     path = os.path.join(user_config_folder, "config.json")
        #
        #     # --- If config is missing, optionally copy default from exe folder ---
        #     if not os.path.exists(path):
        #         default_config_path = os.path.join(os.path.dirname(__file__), "config.json")
        #         if os.path.exists(default_config_path):
        #             shutil.copy(default_config_path, path)
        #         else:
        #             print(f"Default config.json not found. Creating empty config.")
        #             with open(path, 'w', encoding='utf-8') as f:
        #                 json.dump({}, f)
        #
        # # --- Load and return JSON ---
        # try:
        #     with open(path, 'r', encoding='utf-8') as f:
        #         print(json.load(f))
        #         return json.load(f)
        # except Exception as e:
        #     print(f"Error reading config.json: {e}")
        #     return {}
