from dotenv import load_dotenv
import os
import msal
from msal import ConfidentialClientApplication
import requests
import base64
import time
import pandas as pd
from datetime import datetime

load_dotenv()

TENANT_ID = os.getenv("SHAREPOINT_TENANT_ID")
CLIENT_ID = os.getenv("SHAREPOINT_CLIENT_ID")
CLIENT_SECRET = os.getenv("SHAREPOINT_CLIENT_SECRET")
AUTHORITY: str = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["https://graph.microsoft.com/.default"]


class Outlook:
    def __init__(self, sender_email: str = "hub.axt@araxaeng.com.br"):
        """
        Inicializa o cliente Outlook com gerenciamento automático de tokens

        Args:
            sender_email: E-mail remetente (opcional, padrão: hub.axt@araxaeng.com.br)
        """
        self.sender_email = sender_email
        self.access_token = None
        self.token_expires_at = 0

        self.app = msal.ConfidentialClientApplication(
            CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
        )

        self._get_token()

    def _get_token(self) -> None:
        """Obtém ou renova o token de acesso automaticamente"""
        current_time = time.time()

        if not self.access_token or current_time >= (self.token_expires_at - 300):
            result = self.app.acquire_token_for_client(scopes=SCOPES)

            if "access_token" in result:
                self.access_token = result["access_token"]
                self.token_expires_at = current_time + result["expires_in"]
            else:
                error = result.get("error")
                error_desc = result.get("error_description")
                raise Exception(f"Falha ao obter token: {error} - {error_desc}")
        else:
            remaining = int(self.token_expires_at - current_time)

    def enviar_email(
        self,
        subject: str,
        to_recipients: list[str],
        body: str,
        cc_recipients: list[str] = None,
        attachments: list[dict] = None,
    ) -> requests.Response:
        """
        Envia e-mail através do Microsoft Graph API

        Args:
            subject: Assunto do e-mail
            to_recipients: Lista de e-mails dos destinatários principais
            body: Conteúdo HTML do e-mail
            cc_recipients: Lista de e-mails em cópia (opcional)
            attachments: Lista de anexos (opcional)
                Formato: [{"file_name": "nome.txt", "content": b"bytes_do_arquivo"}]

        Returns:
            Objeto Response da requisição
        """
        # Garante token válido
        self._get_token()

        # Construir destinatários
        to_recipients_list = [
            {"emailAddress": {"address": email}} for email in to_recipients
        ]

        # Construir cópias
        cc_recipients_list = []
        if cc_recipients:
            cc_recipients_list = [
                {"emailAddress": {"address": email}} for email in cc_recipients
            ]

        # Payload da mensagem
        message_payload = {
            "subject": subject,
            "body": {"contentType": "HTML", "content": body},
            "toRecipients": to_recipients_list,
            "ccRecipients": cc_recipients_list,
        }

        # Adicionar anexos
        if attachments:
            message_payload["attachments"] = []
            for attachment in attachments:
                message_payload["attachments"].append(
                    {
                        "@odata.type": "#microsoft.graph.fileAttachment",
                        "name": attachment["file_name"],
                        "contentBytes": base64.b64encode(attachment["content"]).decode(
                            "utf-8"
                        ),
                    }
                )

        # Endpoint de envio
        endpoint = (
            f"https://graph.microsoft.com/v1.0/users/{self.sender_email}/sendMail"
        )

        response = requests.post(
            endpoint,
            headers={
                "Authorization": f"Bearer {self.access_token}",
                "Content-Type": "application/json",
            },
            json={"message": message_payload},
        )

        return response


outlook = Outlook()


class Sharepoint:
    @staticmethod
    def obter_token():

        authority = f"https://login.microsoftonline.com/{TENANT_ID}"
        escopo = ["https://graph.microsoft.com/.default"]

        app = ConfidentialClientApplication(
            client_id=CLIENT_ID,
            authority=authority,
            client_credential=CLIENT_SECRET,
        )
        resultado = app.acquire_token_for_client(scopes=escopo)

        token_acesso = resultado.get("access_token")
        if not token_acesso:
            print("Erro ao obter token de acesso:", resultado)

        cabecalhos = {
            "Authorization": f"Bearer {token_acesso}",
            "Accept": "application/json",
        }

        return cabecalhos

    @staticmethod
    def ler_civ(nome: str, e_etapa: str) -> str:
        """
        Lê o item completo da lista CIV pelo id retornado em ler_civ.
        Retorna o JSON do item.
        """

        def ler_civ_por_nome(nomedocumento: str) -> pd.DataFrame:
            list_items_url = "https://graph.microsoft.com/v1.0/sites/araxaenergiasolar.sharepoint.com,14d3a376-d3d1-42fd-b9b4-0dfc433e7239,cdcf2e8b-0d8f-47e1-944e-8a1c538bb229/lists/e1cdd03b-3bd2-4b2a-8eb5-c93591d42e17/items?$filter=fields/Cod_Araxa eq '{}'&$top=1".format(
                nomedocumento
            )
            response = requests.get(list_items_url, headers=Sharepoint.obter_token())
            list_items = response.json().get("value")
            if not list_items:
                raise FileNotFoundError(f"Documento '{nomedocumento}' não encontrado.")
            df = pd.DataFrame(list_items)
            return df

        item_id = ler_civ_por_nome(nome)["id"][0]

        url = f"https://graph.microsoft.com/v1.0/sites/araxaenergiasolar.sharepoint.com,14d3a376-d3d1-42fd-b9b4-0dfc433e7239,cdcf2e8b-0d8f-47e1-944e-8a1c538bb229/lists/e1cdd03b-3bd2-4b2a-8eb5-c93591d42e17/items/{item_id}"
        response = requests.get(url, headers=Sharepoint.obter_token())
        response.raise_for_status()

        if e_etapa == "etapa":
            resposta = response.json()["fields"]["Etapa"]
            return resposta
        elif e_etapa == "disciplina":
            disciplina = response.json()["fields"]["Disciplina_Cod"]
            return disciplina
        elif e_etapa == "revisao":
            revisao = response.json()["fields"]["RevCliente"]
            return revisao


if __name__ == "__main__":
    now = str(datetime.now())
    fim = Sharepoint.ler_civ("7097-LT-C0-MC-0089-0A", "etapa")
    outlook.enviar_email(
        subject="Teste N8N",
        to_recipients=["caio.alves@araxaeng.com.br"],
        cc_recipients=["jorge.diaz@araxaeng.com.br"],
        body=f"""Teste feito às {now} e o status do documento "7097-LT-C0-MC-0089" é {fim}""",
    )
