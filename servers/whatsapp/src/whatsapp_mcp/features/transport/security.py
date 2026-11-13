from __future__ import annotations

from mcp.server.transport_security import TransportSecuritySettings

LOOPBACK_HOSTS = ("127.0.0.1:*", "localhost:*", "[::1]:*")


def transport_security(allowed_hosts: tuple[str, ...]) -> TransportSecuritySettings | None:
    if not allowed_hosts:
        return None
    hosts = _unique(
        [*LOOPBACK_HOSTS, *(pattern for host in allowed_hosts for pattern in _host_patterns(host))]
    )
    origins = _unique([f"{scheme}://{host}" for host in hosts for scheme in ("http", "https")])
    return TransportSecuritySettings(
        enable_dns_rebinding_protection=True,
        allowed_hosts=hosts,
        allowed_origins=origins,
    )


def _host_patterns(host: str) -> tuple[str, ...]:
    if _has_port(host):
        return (host,)
    return (host, f"{host}:*")


def _has_port(host: str) -> bool:
    if host.startswith("["):
        return "]:" in host
    return ":" in host


def _unique(values: list[str]) -> list[str]:
    return list(dict.fromkeys(values))
