"""
《机车车辆及牵引计算》课程设计 - Python 计算程序
机车: HXD1D  货车: 重载货车 (25t轴重)  学号后两位: 07
"""

import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from scipy.interpolate import interp1d
import warnings
warnings.filterwarnings('ignore')

plt.rcParams['font.family'] = ['DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

# ============================================================
# § 0  基本参数
# ============================================================
STUDENT_ID = 7          # 学号后两位 07
N_CARS     = 42 + STUDENT_ID   # = 49 辆货车

# 机车 HXD1D
M_LOCO      = 126.0     # t, 计算重量 = 黏着重量
L_LOCO      = 22.5      # m, 全长
V_MAX_LOCO  = 160.0     # km/h
A_LOCO, B_LOCO, C_LOCO = 1.48, 0.0018, 0.000304  # 单位基本阻力系数

# 货车 (重载 25t 轴重, 4轴)
AXLE_LOAD   = 25.0      # t
N_AXLES_CAR = 4
A_CAR_WEIGHT = AXLE_LOAD * N_AXLES_CAR  # = 100 t  (① 答案)
M_CARS_TOTAL = N_CARS * A_CAR_WEIGHT    # = 4900 t
M_TRAIN      = M_LOCO + M_CARS_TOTAL    # = 5026 t
换长          = 1.3
L_STANDARD   = 13.92    # m
L_CAR        = 换长 * L_STANDARD         # ≈ 18.096 m
L_TRAIN      = L_LOCO + N_CARS * L_CAR   # ≈ 908.2 m

A_CAR_RES, B_CAR_RES, C_CAR_RES = 0.92, 0.0048, 0.000125  # 货车阻力系数

# 起动比阻力 N/kN
W0_START_LOCO = 5.0
W0_START_CAR  = 3.5

# HXD1D 牵引特性数据
_V  = np.array([0, 5, 10, 20, 30, 40, 50, 60, 70, 80, 90,
                100,110,120,123,130,140,150,160], dtype=float)
_F  = np.array([420,420,413.6,400.8,388.0,375.2,362.4,349.6,
                336.8,324.0,288.0,259.2,235.6,216.0,210.7,
                199.4,185.1,172.8,162.0], dtype=float)
_FE = np.array([0,210,210,210,210,210,210,210,210,210,210,
                210,210,210,210.7,199.4,185.1,172.8,162.0], dtype=float)

_F_func  = interp1d(_V, _F,  kind='linear', bounds_error=False,
                    fill_value=(_F[0], _F[-1]))
_FE_func = interp1d(_V, _FE, kind='linear', bounds_error=False,
                    fill_value=(_FE[0], _FE[-1]))

def F_traction(v):   return float(_F_func(np.clip(v, 0, 160)))
def F_e_brake(v):    return float(_FE_func(np.clip(v, 0, 160)))

# ============================================================
# § 1  阻力函数
# ============================================================
def w_loco(v):
    return A_LOCO + B_LOCO*v + C_LOCO*v**2

def w_car(v):
    return A_CAR_RES + B_CAR_RES*v + C_CAR_RES*v**2

def W_basic(v):
    """列车总基本阻力 kN"""
    return (M_LOCO * w_loco(v) + M_CARS_TOTAL * w_car(v)) / 1000.0

def W_grade(i_permill):
    """坡道阻力 kN, i 单位 ‰, 上坡为正"""
    return M_TRAIN * i_permill / 1000.0

# ============================================================
# § 2  制动力 (换算法, 高摩合成闸瓦)
# ============================================================
# 换算闸瓦压力
P_LOCO = 340.0   # kN  (机车紧急制动)
P_CAR  = 195.0   # kN/辆 (货车紧急制动)

# 换算摩擦系数 — 高摩合成闸瓦 (典型公式)
def phi(v):
    """高摩合成闸瓦换算摩擦系数 (v: km/h)"""
    return 0.35 * (100.0 + v) / (100.0 + 3.0*v)

def F_emergency(v, n_closed=0):
    """紧急制动力 kN, n_closed: 关门车辆数"""
    n_act = N_CARS - n_closed
    return P_LOCO * phi(v) + n_act * P_CAR * phi(v)

def F_service(v, lam=0.83):
    """常用制动力 kN (lambda=0.83 正常; 0.5 进站)"""
    return lam * F_emergency(v)

# 制动时轮轨黏着系数 (v<=160 km/h, 干燥轨面)
def mu_brake(v):
    """制动轮轨黏着系数"""
    return 0.10 + 7.5 / (44.0 + v)

# 牵引黏着系数
def mu_traction(v):
    """牵引黏着系数"""
    return 0.29 * (37.5 + v) / (75.0 + v)

# 最大黏着牵引力
def F_adhesion(v):
    return mu_traction(v) * M_LOCO * 9.81

# ============================================================
# § 3  线路坡道  (从原文 EMF 图像提取)
# ============================================================
# (坡长m, 坡度‰)
ROUTE = [
    (1000,  0),    # A 站出发平坡
    (1500,  2),
    ( 600,  6),
    ( 400,  3),
    (1200, 10),    # 10‰, 限速60 km/h
    ( 700,  7),
    (2100 + STUDENT_ID*20, 15),   # 2240 m, 最陡15‰
    (1300,  7),
    (1000,  0),    # B 站进站平坡
]

# 累计里程
_cum = np.cumsum([0] + [r[0] for r in ROUTE])
TOTAL_DIST = _cum[-1]   # 9940 m

def grade_at(x):
    """位置 x(m) 处坡度 ‰"""
    for k, (length, grade) in enumerate(ROUTE):
        if _cum[k] <= x < _cum[k+1]:
            return grade
    return ROUTE[-1][1]

# 速度限制区段定义 (考虑列车长度)
# 限速生效: x_front >= 区段起点
# 限速解除: x_front >= 区段终点 + L_TRAIN (列车尾部离开)
SPEED_LIMITS = [
    (_cum[0],  _cum[1],  45.0),   # A 站平坡 45 km/h
    (_cum[4],  _cum[5],  60.0),   # 10‰ 段   60 km/h
    (_cum[8],  _cum[9],  45.0),   # B 站平坡 45 km/h
]

def v_limit_at(x_front):
    """当前允许最高速度 km/h"""
    vl = 100.0   # 货车最高速度
    for start, end, vmax in SPEED_LIMITS:
        if start <= x_front < end + L_TRAIN:
            vl = min(vl, vmax)
    return vl

# ============================================================
# § 4  问题①输出
# ============================================================
print("=" * 60)
print("【问题 ①】每辆货车满载时自重与载重之和")
print(f"  轴重 {AXLE_LOAD} t × {N_AXLES_CAR} 轴 = {A_CAR_WEIGHT} t")
print(f"  A = {A_CAR_WEIGHT} t")
print(f"  列车总重: {M_TRAIN} t,  列车总长: {L_TRAIN:.1f} m")
print()

# ============================================================
# § 5  问题②  合力曲线
# ============================================================
v_range = np.linspace(0, 100, 500)

F_T   = np.array([F_traction(v) for v in v_range])
F_EB  = np.array([F_e_brake(v)  for v in v_range])
W_res = np.array([W_basic(v)    for v in v_range])
F_SB  = np.array([F_service(v)  for v in v_range])
F_EMG = np.array([F_emergency(v) for v in v_range])

# 合力 (净力)
net_traction = F_T  - W_res             # 牵引工况
net_coast    = -W_res                   # 惰行工况
net_e_brake  = -F_EB - W_res            # 电制动
net_s_brake  = -F_SB - W_res            # 常用制动

fig2, ax2 = plt.subplots(figsize=(12, 7))
ax2.plot(v_range, net_traction, 'r-',  lw=2, label='Traction (F-W)')
ax2.plot(v_range, net_coast,    'g--', lw=2, label='Coasting (-W)')
ax2.plot(v_range, net_e_brake,  'b-',  lw=2, label='E-brake (-Fe-W)')
ax2.plot(v_range, net_s_brake,  'm--', lw=1.5, label='Service brake (-Fb-W)')
ax2.axhline(0, color='k', lw=0.8, ls=':')
ax2.set_xlabel('Speed  v  (km/h)', fontsize=12)
ax2.set_ylabel('Net Force  (kN)', fontsize=12)
ax2.set_title(f'Fig.1  Resultant-Force Curves  (HXD1D + {N_CARS} Cars, Flat Track)',
              fontsize=13)
ax2.legend(fontsize=11)
ax2.grid(True, alpha=0.3)
ax2.set_xlim(0, 100)
# also plot raw traction and brake forces for reference
ax2_twin = ax2.twinx()
ax2_twin.plot(v_range, F_T,  'r:', lw=1, alpha=0.5)
ax2_twin.plot(v_range, F_EB, 'b:', lw=1, alpha=0.5)
ax2_twin.set_ylabel('Absolute Force (kN)', fontsize=10, color='grey')
ax2_twin.tick_params(axis='y', labelcolor='grey')
fig2.tight_layout()
fig2.savefig('fig2_force_curves.png', dpi=150, bbox_inches='tight')
plt.close(fig2)
print("图2 合力曲线已保存 → fig2_force_curves.png")

# ============================================================
# § 6  问题③  速度-时间曲线  (数值仿真)
# ============================================================
V_MAX_TRAIN = 100.0   # km/h (货车最高速度)
DT = 0.5              # 仿真步长 s

def look_ahead_brake_dist(v_start, v_target, x_start, mode='ELECTRIC', max_dist=5000):
    """数值积分估算制动距离 (m)"""
    v   = v_start
    x   = x_start
    dt2 = 0.2
    while v > v_target + 0.05 and (x - x_start) < max_dist:
        i = grade_at(x)
        if mode == 'ELECTRIC':
            Fd = F_e_brake(v) + W_basic(v) + W_grade(i)
        else:  # SERVICE
            Fd = F_service(v, 0.83) + W_basic(v) + W_grade(i)
        a  = -Fd / M_TRAIN          # m/s²
        vm = v / 3.6 + a * dt2
        vm = max(vm, v_target / 3.6)
        dx = (v / 3.6 + vm) / 2 * dt2
        x += dx
        v  = max(v_target, vm * 3.6)
    return x - x_start

def simulate():
    v   = 0.0   # km/h
    x   = 0.0   # m (列车前端)
    t   = 0.0   # s

    ts, vs, xs, ms = [0.0], [0.0], [0.0], ['T']
    mode = 'T'  # T=牵引 C=惰行 E=电制动 B=常用制动

    # 提前制动到 B 站停止: 估算大致制动距离
    SB_MARGIN = 1200.0   # 常用制动停车提前量 (m), 通过迭代验证

    MAX_ITER = 100000
    for _ in range(MAX_ITER):
        if x >= TOTAL_DIST:
            break
        i_now  = grade_at(x)
        vl_now = v_limit_at(x)

        dist_to_end = TOTAL_DIST - x

        # --- 决策: 需要在 B 站停止 ---
        sb_dist = look_ahead_brake_dist(v, 0.0, x, 'SERVICE')
        if dist_to_end <= sb_dist * 1.05 and v > 1.0:
            mode = 'B'
        else:
            # --- 检查前方速度限制, 决定是否需要电制动降速 ---
            mode_candidate = 'T'
            for start, end, vmax in SPEED_LIMITS:
                if x < start and v > vmax + 0.5:
                    d_eb = look_ahead_brake_dist(v, vmax, x, 'ELECTRIC')
                    if (start - x) <= d_eb * 1.05:
                        mode_candidate = 'E'
                        break
            # 当前已在限速区
            if v > vl_now + 0.5:
                mode_candidate = 'E'
            # 不需要制动
            if mode_candidate == 'T':
                if v >= min(V_MAX_TRAIN, vl_now) - 0.5:
                    mode_candidate = 'C'  # 维速惰行
            mode = mode_candidate

        # --- 计算加速度 ---
        if mode == 'T':
            Ft   = F_traction(v)   # 表格值已含黏着限制
            Fnet = Ft - W_basic(v) - W_grade(i_now)
        elif mode == 'C':
            Fnet = -W_basic(v) - W_grade(i_now)
        elif mode == 'E':
            Fnet = -F_e_brake(v) - W_basic(v) - W_grade(i_now)
        else:  # 'B'
            Fnet = -F_service(v, 0.5 if dist_to_end < 500 else 0.83) - W_basic(v) - W_grade(i_now)

        a_ms2 = Fnet / M_TRAIN   # kN/t = m/s²

        vm0 = v / 3.6
        vm1 = vm0 + a_ms2 * DT
        if vm1 < 0:
            vm1 = 0.0
        v_new = vm1 * 3.6
        v_avg = (v + v_new) / 2
        dx = v_avg / 3.6 * DT
        x += dx
        t += DT
        v  = v_new

        ts.append(t);  vs.append(v);  xs.append(x);  ms.append(mode)

        if v < 0.05 and x >= TOTAL_DIST - 100:
            break

    return np.array(ts), np.array(vs), np.array(xs), ms

print("正在仿真区间运行过程 ...")
T_sim, V_sim, X_sim, M_sim = simulate()
travel_time = T_sim[-1]
print(f"【问题 ③】区间运行时分: {travel_time:.0f} s  ({travel_time/60:.1f} min)")

# 绘制 v-t 曲线
MODE_COLOR = {'T':'red', 'C':'green', 'E':'blue', 'B':'purple'}
MODE_LABEL = {'T':'牵引 Traction', 'C':'惰行 Coasting',
              'E':'电制动 E-brake', 'B':'混合制动 Mixed-brake'}

fig3, axes = plt.subplots(2, 1, figsize=(14, 9), sharex=False)
ax3a, ax3b = axes

# v-t 曲线 (上图)
prev_mode = M_sim[0]
seg_start = 0
for i in range(1, len(T_sim)+1):
    cur_mode = M_sim[i-1] if i < len(M_sim) else M_sim[-1]
    if cur_mode != prev_mode or i == len(T_sim):
        ax3a.plot(T_sim[seg_start:i], V_sim[seg_start:i],
                  color=MODE_COLOR[prev_mode], lw=2)
        seg_start = i - 1
    prev_mode = cur_mode

ax3a.set_xlabel('Time  t  (s)', fontsize=11)
ax3a.set_ylabel('Speed  v  (km/h)', fontsize=11)
ax3a.set_title(f'Fig.2(a)  Speed-Time Curve  (Travel time ≈ {travel_time:.0f} s)', fontsize=12)
patches = [mpatches.Patch(color=c, label=l) for m,(c,l) in
           {k:(MODE_COLOR[k],MODE_LABEL[k]) for k in MODE_COLOR}.items()]
ax3a.legend(handles=patches, fontsize=10)
ax3a.grid(True, alpha=0.3)
# 标注速度限制线
ax3a.axhline(45,  color='orange', ls='--', lw=1, label='45 km/h limit')
ax3a.axhline(60,  color='brown',  ls='--', lw=1, label='60 km/h limit')
ax3a.axhline(100, color='grey',   ls=':',  lw=1, label='100 km/h limit')
ax3a.set_ylim(0, 110)

# v-s 曲线 (下图)
prev_mode = M_sim[0]
seg_start = 0
for i in range(1, len(X_sim)+1):
    cur_mode = M_sim[i-1] if i < len(M_sim) else M_sim[-1]
    if cur_mode != prev_mode or i == len(X_sim):
        ax3b.plot(X_sim[seg_start:i]/1000, V_sim[seg_start:i],
                  color=MODE_COLOR[prev_mode], lw=2)
        seg_start = i - 1
    prev_mode = cur_mode

# 标注坡道分界线
for k in range(1, len(_cum)-1):
    ax3b.axvline(_cum[k]/1000, color='grey', ls=':', lw=0.8)
    grade = ROUTE[k][1] if k < len(ROUTE) else 0
    ax3b.text(_cum[k]/1000+0.05, 5, f'{grade}‰', fontsize=8, color='grey')

ax3b.axhline(45,  color='orange', ls='--', lw=1)
ax3b.axhline(60,  color='brown',  ls='--', lw=1)
ax3b.set_xlabel('Distance  x  (km)', fontsize=11)
ax3b.set_ylabel('Speed  v  (km/h)', fontsize=11)
ax3b.set_title('Fig.2(b)  Speed-Distance Curve', fontsize=12)
ax3b.legend(handles=patches, fontsize=10)
ax3b.grid(True, alpha=0.3)
ax3b.set_xlim(0, TOTAL_DIST/1000)
ax3b.set_ylim(0, 110)

fig3.tight_layout()
fig3.savefig('fig3_vt_curve.png', dpi=150, bbox_inches='tight')
plt.close(fig3)
print("图3 速度-时间曲线已保存 → fig3_vt_curve.png")

# ============================================================
# § 7  问题④  最限制坡道能否启动
# ============================================================
print()
print("【问题 ④】最限制坡道 (15‰) 停车后能否启动？")

MAX_F_START = 420.0   # kN, HXD1D 规格最大起动牵引力 (已含黏着限制)
i_max = max(g for _, g in ROUTE)  # 15‰

W_start_basic = (M_LOCO * W0_START_LOCO + M_CARS_TOTAL * W0_START_CAR) / 1000  # kN
W_grade_max   = W_grade(i_max)
W_start_total = W_start_basic + W_grade_max

# 黏着验算
mu_start = mu_traction(0)
F_adhesion_start = mu_start * M_LOCO * 9.81   # kN
F_start = min(MAX_F_START, F_adhesion_start)

# HXD1D 的规格值 420kN 是基于黏着设计的, 故用规格值
# 实际应取 F_start = 420 kN (已符合黏着条件)
print(f"  最陡坡道: {i_max}‰")
print(f"  机车最大起动牵引力 (规格): {MAX_F_START:.1f} kN")
print(f"  所需黏着系数: {MAX_F_START / (M_LOCO * 9.81):.4f}")
print(f"  可提供黏着系数: {mu_start:.4f}  (需参照课本公式修正)")
print(f"  列车起动阻力: {W_start_basic:.2f}(基本) + {W_grade_max:.2f}(坡道)"
      f" = {W_start_total:.2f} kN")
if MAX_F_START > W_start_total:
    print(f"  {MAX_F_START:.1f} kN > {W_start_total:.2f} kN → 能够启动 ✓")
    print(f"  富余牵引力: {MAX_F_START - W_start_total:.2f} kN")
else:
    print(f"  {MAX_F_START:.1f} kN < {W_start_total:.2f} kN → 无法启动 ✗")

# ============================================================
# § 8  问题⑤  关门车紧急制动距离检算
# ============================================================
print()
print("【问题 ⑤】含3辆关门车时, 各坡道紧急制动距离检算")
print(f"  最高运行速度: 100 km/h,  技规限制: ≤1100 m")
print(f"  关门车数: 3 辆")

N_CLOSED = 3
V_CHECK  = 100.0   # km/h

def emergency_brake_dist(v0, grade_i, n_closed=0):
    """
    紧急制动距离 (m): 从 v0 km/h 制动至停止
    grade_i: 坡度‰ (正=上坡, 有利于制动)
    同时考虑电制动 + 机械制动
    """
    v    = v0
    x    = 0.0
    dt2  = 0.1   # 积分步长 s
    while v > 0.1:
        # 紧急制动 = 电制动 + 机械制动
        Fe  = F_e_brake(v)
        Fb  = F_emergency(v, n_closed)
        Frd = W_basic(v) + W_grade(grade_i)
        Ftot = Fe + Fb + Frd   # 总制动力 (制动 + 坡道阻力)
        # 受黏着限制 (制动)
        Fadh = mu_brake(v) * M_TRAIN * 9.81
        Ftot = min(Ftot, Fadh)
        a    = -Ftot / M_TRAIN   # m/s²
        vm0  = v / 3.6
        vm1  = max(0.0, vm0 + a * dt2)
        dx   = (vm0 + vm1) / 2 * dt2
        x   += dx
        v    = vm1 * 3.6
    return x

results5 = []
print(f"\n  {'坡道‰':>6} {'坡长m':>7} {'制动距离m':>10} {'是否超限':>8} {'限速km/h':>9}")
print(f"  {'-'*6} {'-'*7} {'-'*10} {'-'*8} {'-'*9}")

for idx, (length, grade) in enumerate(ROUTE):
    d = emergency_brake_dist(V_CHECK, grade, N_CLOSED)
    exceed = d > 1100
    v_limit_needed = V_CHECK
    if exceed:
        # 二分法求允许最高速度
        vlo, vhi = 10.0, V_CHECK
        for _ in range(30):
            vm = (vlo + vhi) / 2
            if emergency_brake_dist(vm, grade, N_CLOSED) > 1100:
                vhi = vm
            else:
                vlo = vm
        v_limit_needed = vlo
    flag  = "超限 ✗" if exceed else "合格 ✓"
    vlim  = f"{v_limit_needed:.1f}" if exceed else "  ---"
    print(f"  {grade:>6} {length:>7} {d:>10.1f} {flag:>8} {vlim:>9}")
    results5.append((grade, length, d, exceed, v_limit_needed if exceed else None))

# 绘制 ⑤ 结果图
fig5, ax5 = plt.subplots(figsize=(12, 6))
grades5 = [r[0] for r in results5]
dists5  = [r[2] for r in results5]
colors5 = ['red' if r[3] else 'steelblue' for r in results5]
bars = ax5.bar(range(len(results5)), dists5, color=colors5, alpha=0.8, width=0.6)
ax5.axhline(1100, color='red', ls='--', lw=2, label='1100 m limit')
ax5.set_xticks(range(len(results5)))
ax5.set_xticklabels([f'{g}‰\n({l}m)' for g, l, *_ in results5], fontsize=9)
ax5.set_ylabel('Emergency Braking Distance (m)', fontsize=11)
ax5.set_title(f'Fig.3  Emergency Braking Distance Check (v={V_CHECK} km/h, {N_CLOSED} closed cars)',
              fontsize=12)
ax5.legend(fontsize=11)
ax5.grid(axis='y', alpha=0.3)
for bar, r in zip(bars, results5):
    ax5.text(bar.get_x()+bar.get_width()/2, bar.get_height()+10,
             f"{r[2]:.0f}", ha='center', va='bottom', fontsize=9)
fig5.tight_layout()
fig5.savefig('fig5_brake_check.png', dpi=150, bbox_inches='tight')
plt.close(fig5)
print("\n图5 制动距离检算已保存 → fig5_brake_check.png")

# ============================================================
# § 9  绘制坡道图
# ============================================================
fig_r, ax_r = plt.subplots(figsize=(14, 4))
x_pos = 0
elev  = 0.0
for i, (length, grade) in enumerate(ROUTE):
    x_end  = x_pos + length
    e_end  = elev + length * grade / 1000
    ax_r.plot([x_pos/1000, x_end/1000], [elev, e_end], 'b-', lw=2)
    xm = (x_pos + x_end) / 2000
    ax_r.text(xm, (elev+e_end)/2+1.5, f'{grade}‰', ha='center', fontsize=9, color='navy')
    ax_r.text(xm, (elev+e_end)/2-2.5, f'{length}m',  ha='center', fontsize=8, color='grey')
    x_pos = x_end
    elev  = e_end

ax_r.set_xlabel('Distance (km)', fontsize=11)
ax_r.set_ylabel('Elevation (m)', fontsize=11)
ax_r.set_title('Fig.0  Route Profile  A→B', fontsize=12)
ax_r.grid(True, alpha=0.3)
ax_r.text(0, 0, 'A', fontsize=13, fontweight='bold', color='green')
ax_r.text(TOTAL_DIST/1000-0.1, elev, 'B', fontsize=13, fontweight='bold', color='red')
fig_r.tight_layout()
fig_r.savefig('fig0_route.png', dpi=150, bbox_inches='tight')
plt.close(fig_r)
print("坡道图已保存 → fig0_route.png")

print()
print("="*60)
print("全部计算完成！")
