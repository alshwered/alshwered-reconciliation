# Alshwered Reconciliation

A Streamlit app for reconciling Zid order exports against normalized payment-provider transaction exports.

## Run locally

```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\\Scripts\\activate
pip install -r requirements.txt
streamlit run app.py
```

## Deploy

Push this folder to a private GitHub repository, then deploy `app.py` on Streamlit Community Cloud.

## Security

Never commit API credentials, `.env`, or `.streamlit/secrets.toml`. Rotate any credentials that were shared in chat. Add new credentials only to the hosting provider's Secrets manager after provider-specific API permissions and endpoints are verified.

## Current scope

- Reads Zid Excel exports and collapses continuation product rows to one order row.
- Normalizes Saudi mobile numbers.
- Matches payment exports by order reference first, then mobile + amount.
- Flags likely duplicate invoices for human review.
- Filters and exports reconciliation results to Excel.
- Direct API settlement adapters are intentionally not enabled until the merchant reporting endpoints and scopes are verified for each provider.
