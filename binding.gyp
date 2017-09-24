{
  "target": [
    {
      "target_name": "talib-binding",
      "sources": [
        "./src/talib-binding.generated.cc"
      ],
      "include_dirs": [
        "<!(node -e \"require('nan')\")"
      ],
      "link_settings": {
        "libraries": [
          "../src/lib/lib/libta_libc_csr.a"
        ]
      }
    }
  ]
}