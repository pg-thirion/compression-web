#!/usr/bin/env bash
set -euo pipefail

REGION="${AWS_REGION:-eu-west-3}"
ACCOUNT_ID="${AWS_ACCOUNT_ID:-$(aws sts get-caller-identity --query Account --output text)}"
BUCKET_NAME="${BUCKET_NAME:-hbtrained-compression-web-${ACCOUNT_ID}}"
SCRIPT_DIR="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"

if ! aws s3api head-bucket --bucket "$BUCKET_NAME" --region "$REGION" >/dev/null 2>&1; then
  aws s3api create-bucket \
    --bucket "$BUCKET_NAME" \
    --region "$REGION" \
    --create-bucket-configuration LocationConstraint="$REGION"
fi

aws s3api put-public-access-block \
  --bucket "$BUCKET_NAME" \
  --region "$REGION" \
  --public-access-block-configuration \
  BlockPublicAcls=true,IgnorePublicAcls=true,BlockPublicPolicy=false,RestrictPublicBuckets=false

aws s3api put-bucket-website \
  --bucket "$BUCKET_NAME" \
  --region "$REGION" \
  --website-configuration '{"IndexDocument":{"Suffix":"index.html"}}'

POLICY_JSON=$(cat <<JSON
{
  "Version": "2012-10-17",
  "Statement": [
    {
      "Sid": "PublicReadSite",
      "Effect": "Allow",
      "Principal": "*",
      "Action": "s3:GetObject",
      "Resource": [
        "arn:aws:s3:::${BUCKET_NAME}/index.html",
        "arn:aws:s3:::${BUCKET_NAME}/favicon.png",
        "arn:aws:s3:::${BUCKET_NAME}/favicon.ico",
        "arn:aws:s3:::${BUCKET_NAME}/app.js",
        "arn:aws:s3:::${BUCKET_NAME}/lib/jszip.min.js",
        "arn:aws:s3:::${BUCKET_NAME}/lib/pdf-lib.min.js"
      ]
    }
  ]
}
JSON
)

aws s3api put-bucket-policy \
  --bucket "$BUCKET_NAME" \
  --region "$REGION" \
  --policy "$POLICY_JSON"

aws s3 cp "$SCRIPT_DIR/index.html" "s3://${BUCKET_NAME}/index.html" --region "$REGION" --content-type "text/html" --cache-control "no-cache"
aws s3 cp "$SCRIPT_DIR/favicon.png" "s3://${BUCKET_NAME}/favicon.png" --region "$REGION" --content-type "image/png" --cache-control "public, max-age=86400"
aws s3 cp "$SCRIPT_DIR/favicon.ico" "s3://${BUCKET_NAME}/favicon.ico" --region "$REGION" --content-type "image/x-icon" --cache-control "public, max-age=86400"
aws s3 cp "$SCRIPT_DIR/app.js" "s3://${BUCKET_NAME}/app.js" --region "$REGION" --content-type "text/javascript" --cache-control "public, max-age=86400"
aws s3 cp "$SCRIPT_DIR/lib/jszip.min.js" "s3://${BUCKET_NAME}/lib/jszip.min.js" --region "$REGION" --content-type "text/javascript" --cache-control "public, max-age=31536000, immutable"
aws s3 cp "$SCRIPT_DIR/lib/pdf-lib.min.js" "s3://${BUCKET_NAME}/lib/pdf-lib.min.js" --region "$REGION" --content-type "text/javascript" --cache-control "public, max-age=31536000, immutable"

CLOUDFRONT_DISTRIBUTION_ID="${CLOUDFRONT_DISTRIBUTION_ID:-}"
if [ -z "$CLOUDFRONT_DISTRIBUTION_ID" ] && [ -f "$HOME/.square-fit-batch/state" ]; then
  CLOUDFRONT_DISTRIBUTION_ID=$(head -n 1 "$HOME/.square-fit-batch/state")
fi
if [ -n "$CLOUDFRONT_DISTRIBUTION_ID" ]; then
  INVALIDATION_BATCH=".invalidation-batch.json"
  cat > "$INVALIDATION_BATCH" <<JSON
{
  "Paths": {
    "Quantity": 6,
    "Items": ["/index.html", "/favicon.png", "/favicon.ico", "/app.js", "/lib/jszip.min.js", "/lib/pdf-lib.min.js"]
  },
  "CallerReference": "deploy-$(date +%s)"
}
JSON
  aws cloudfront create-invalidation \
    --distribution-id "$CLOUDFRONT_DISTRIBUTION_ID" \
    --invalidation-batch "file://${INVALIDATION_BATCH}" \
    --output json >/dev/null && echo "Invalidated CloudFront cache for index.html, app.js, lib/*"
  rm -f "$INVALIDATION_BATCH"
fi

DIST_DOMAIN=""
if [ -f "$HOME/.square-fit-batch/state" ]; then
  DIST_DOMAIN=$(sed -n '2p' "$HOME/.square-fit-batch/state")
fi
if [ -n "$DIST_DOMAIN" ]; then
  printf 'HTTPS: https://%s\n' "$DIST_DOMAIN"
fi
printf 'HTTP (fallback): http://%s.s3-website.%s.amazonaws.com\n' "$BUCKET_NAME" "$REGION"
