"""
backend_connector.py
Gestion des connexions Ollama et LM Studio
Station UX Audit v2.6.4
"""

import os
import json
import re
import requests
import subprocess
import time
from config_prompts import reparer_json


class BackendConnector:
    """Gère les connexions avec les backends IA (Ollama, LM Studio)."""

    def __init__(self, log_callback=None):
        self.log_callback = log_callback
        self.base_url = "http://localhost:11434"
        self.backend = "ollama"
        self.modele = ""

    def set_config(self, url: str, backend: str, modele: str):
        """Configure l'URL, le backend et le modèle."""
        self.base_url = url.rstrip("/")
        self.backend = backend
        self.modele = modele

    def _log(self, texte: str):
        """Log via callback si disponible."""
        if self.log_callback:
            self.log_callback(texte)

    def tester_serveur(self) -> list:
        """
        Teste la connexion avec le serveur et retourne la liste des modèles.
        Retourne [] en cas d'erreur.
        """
        try:
            if self.backend == "lmstudio":
                r = requests.get(f"{self.base_url}/v1/models", timeout=4)
                r.raise_for_status()
                modeles = [m["id"] for m in r.json().get("data", [])]
                label = "LM Studio"
            else:
                # Ollama
                r = requests.get(f"{self.base_url}/api/tags", timeout=4)
                r.raise_for_status()
                modeles = [m["name"] for m in r.json().get("models", [])]
                label = "Ollama"

            self._log(f"✅ {label} OK — {len(modeles)} modèle(s) : {', '.join(modeles[:4])}")
            return modeles

        except requests.exceptions.ConnectionError:
            self._log(f"❌ Serveur non joignable sur {self.base_url}")
            return []
        except Exception as e:
            self._log(f"❌ Erreur serveur : {e}")
            return []

    def appel_ia(self, prompt: str, json_force: bool = True) -> dict | None:
        """
        Appelle l'IA pour générer une réponse.
        Retourne un dictionnaire parsé ou None en cas d'erreur.
        """
        if not self.modele:
            self._log("❌ Aucun modèle sélectionné.")
            return None

        brut = ""
        try:
            if self.backend == "lmstudio":
                return self._appel_lmstudio(prompt)
            else:
                return self._appel_ollama(prompt, json_force)

        except Exception as e:
            self._log(f"❌ Erreur IA : {e}")
            return None

    def _appel_ollama(self, prompt: str, json_force: bool = True) -> dict | None:
        """Appel spécifique pour Ollama."""
        payload = {
            "model":   self.modele,
            "prompt":  prompt,
            "stream":  False,
            "options": {"temperature": 0.1, "num_ctx": 262144}
        }
        if json_force:
            payload["format"] = "json"

        try:
            r = requests.post(f"{self.base_url}/api/generate", json=payload, timeout=None)
            if r.status_code != 200:
                self._log(f"❌ Ollama HTTP {r.status_code}")
                return None
            brut = r.json().get("response", "")

            if not brut.strip():
                self._log("❌ Réponse vide")
                return None

            return self._parser_json(brut)

        except Exception as e:
            self._log(f"❌ Ollama : {e}")
            return None

    def _appel_lmstudio(self, prompt: str) -> dict | None:
        """Appel spécifique pour LM Studio."""
        sys_msg = (
            "Tu réponds UNIQUEMENT avec un objet JSON valide. "
            "Aucun texte avant, aucun texte après, aucun bloc markdown, "
            "aucun bloc <think>. Les guillemets dans les valeurs doivent "
            "être remplacés par des guillemets simples '. "
            "Commence directement par '{' et termine par '}'."
        )
        messages = [
            {"role": "system", "content": sys_msg},
            {"role": "user",   "content": prompt}
        ]
        payload = {
            "model":      self.modele,
            "messages":   messages,
            "temperature": 0.1,
            "stream":     False,
            "max_tokens": 262144,
        }

        mod_lower = self.modele.lower()
        if any(k in mod_lower for k in ("qwen3", "qwen3.5", "qwq", "deepseek-r", "gemma-4", "gemma4")):
            payload["enable_thinking"] = False

        headers = {"Connection": "close"}

        try:
            r = requests.post(f"{self.base_url}/v1/chat/completions",
                              json=payload, timeout=None, headers=headers)

            if r.status_code == 400:
                err = r.text
                self._log(f"⚠️  HTTP 400, retry sans system role...")
                merged = f"[CONSIGNE]: {sys_msg}\n\n{prompt}"
                payload["messages"] = [{"role": "user", "content": merged}]
                r = requests.post(f"{self.base_url}/v1/chat/completions",
                                  json=payload, timeout=None, headers=headers)
                if r.status_code != 200:
                    self._log(f"❌ Retry échoué HTTP {r.status_code}")
                    return None
            elif r.status_code != 200:
                self._log(f"❌ LM Studio HTTP {r.status_code}")
                return None

            resp_json = r.json()
            choice    = resp_json.get("choices", [{}])[0]
            msg       = choice.get("message", {})
            brut      = msg.get("content", "") or ""

            if not brut.strip():
                brut = msg.get("reasoning_content", "") or ""
                if brut.strip():
                    self._log("ℹ️  Contenu depuis reasoning_content")

            if not brut.strip():
                self._log(f"❌ Réponse vide")
                return None

            return self._parser_json(brut)

        except Exception as e:
            self._log(f"❌ LM Studio : {e}")
            return None

    def _parser_json(self, brut: str) -> dict | None:
        """Parse le JSON avec tentatives de réparation."""
        # Tentative 1 : parse direct après nettoyage basique
        propre = re.sub(r"<think>.*?</think>", "", brut, flags=re.DOTALL)
        propre = re.sub(r"```json\s*|```\s*", "", propre).strip()
        idx_s  = propre.find("{")
        idx_e  = propre.rfind("}")
        if idx_s != -1 and idx_e != -1:
            extrait = propre[idx_s:idx_e + 1]
            try:
                return json.loads(extrait)
            except json.JSONDecodeError:
                pass

        # Tentative 2 : réparation automatique du JSON
        self._log("🔧 JSON invalide → tentative de réparation...")
        repare = reparer_json(brut)
        try:
            resultat = json.loads(repare)
            self._log("✅ JSON réparé avec succès")
            return resultat
        except json.JSONDecodeError as e2:
            self._log(f"⚠️  Réparation insuffisante : {e2}")
            return None

    def tester_mindmanager(self) -> bool:
        """Teste la connexion avec MindManager (Windows uniquement)."""
        try:
            import win32com.client
            import pythoncom
            pythoncom.CoInitialize()
            try:
                win32com.client.GetActiveObject("MindManager.Application")
                self._log("✅ MindManager connecté.")
                pythoncom.CoUninitialize()
                return True
            except Exception:
                mm = win32com.client.Dispatch("MindManager.Application")
                mm.Visible = True
                self._log("✅ MindManager lancé.")
                pythoncom.CoUninitialize()
                return True
        except ImportError:
            self._log("❌ pywin32 non installé (Windows uniquement)")
            return False
        except Exception as e:
            self._log(f"❌ MindManager : {e}")
            return False

    def optimiser_gpu_amd(self) -> bool:
        """Optimise le GPU AMD Radeon pour Ollama."""
        try:
            import winreg
            variables = {
                "HSA_OVERRIDE_GFX_VERSION": "10.3.0",
                "OLLAMA_NUM_GPU": "1",
                "HIP_VISIBLE_DEVICES": "0",
                "OLLAMA_GPU_OVERHEAD": "0"
            }
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Environment", 0, winreg.KEY_SET_VALUE)
            for n, v in variables.items():
                winreg.SetValueEx(key, n, 0, winreg.REG_SZ, v)
                self._log(f"✅ {n} = {v}")
            winreg.CloseKey(key)

            for n, v in variables.items():
                os.environ[n] = v

            self._redemarrer_ollama()
            return True
        except Exception as e:
            self._log(f"❌ Optimisation GPU : {e}")
            return False

    def _redemarrer_ollama(self):
        """Redémarre le service Ollama."""
        self._log("♻️  Redémarrage Ollama...")
        try:
            subprocess.run(["taskkill", "/F", "/IM", "ollama.exe"], capture_output=True)
            time.sleep(2)
            subprocess.Popen(["ollama", "serve"], creationflags=subprocess.CREATE_NO_WINDOW,
                           env=os.environ.copy())
            self._log("✅ Ollama relancé.")
            time.sleep(3)
        except Exception as e:
            self._log(f"❌ Redémarrage Ollama : {e}")

    def diagnostiquer_gpu(self) -> str:
        """Diagnostique l'état du GPU."""
        try:
            r = requests.get(self.base_url.rstrip("/") + "/api/ps", timeout=5)
            if r.status_code == 200:
                modeles = r.json().get("models", [])
                if not modeles:
                    return "Aucun modèle chargé."
                result = []
                for m in modeles:
                    vram = m.get("size_vram", 0)
                    ram  = max(m.get("size", 0) - vram, 0)
                    result.append(f"{m.get('name','?')} — VRAM: {vram/1e9:.2f}Go | RAM: {ram/1e9:.2f}Go")
                    result.append("✅ GPU ACTIF" if vram > 0 else "⚠️ VRAM=0 — CPU uniquement")
                return "\n".join(result)
        except Exception as e:
            return f"Erreur : {e}"

        return "Impossible de diagnostiquer."
