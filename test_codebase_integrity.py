import os
import sys

sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import app, db, Employee

def test_routes():
    app.config['TESTING'] = True
    client = app.test_client()

    with app.app_context():
        # 1. Test Login GET
        r = client.get('/login')
        assert r.status_code == 200, f"Login GET failed: {r.status_code}"
        print("1. Login GET: OK")

        # 2. Test Admin Dashboard with TDH, A11A, LENS
        with client.session_transaction() as sess:
            sess['employee_code'] = 'ADMIN01'
            sess['role'] = 'admin'
            sess['department'] = 'TDH'

        for dept in ['TDH', 'A11A', 'LENS']:
            r = client.get(f'/projects?dept={dept}')
            assert r.status_code == 200, f"Admin projects ({dept}) failed: {r.status_code}"
            assert b'WORKFLOW' in r.data
            print(f"2. Admin Dashboard ({dept}): OK")

        # 3. Test Admin Employees page
        r = client.get('/admin/employees?dept=TDH')
        assert r.status_code == 200, f"Admin employees failed: {r.status_code}"
        print("3. Admin Employees: OK")

        # 4. Test Leave Calendar page
        r = client.get('/leave-calendar?dept=TDH')
        assert r.status_code == 200, f"Leave calendar failed: {r.status_code}"
        print("4. Leave Calendar: OK")

        # 5. Test User Tasks page
        with client.session_transaction() as sess:
            sess['employee_code'] = 'NVT-01'
            sess['role'] = 'user'
            sess['department'] = 'TDH'

        r = client.get('/user/tasks')
        assert r.status_code == 200, f"User tasks failed: {r.status_code}"
        print("5. User Tasks: OK")

        # 6. Test Excel Export route
        r = client.get('/export/employee/NVT-01')
        assert r.status_code == 200, f"Export failed: {r.status_code}"
        cd = r.headers.get('Content-Disposition', '')
        assert 'BaoCaoCaNhan_' in cd, f"Unexpected export filename: {cd}"
        print(f"6. Excel Export (NVT-01): OK ({cd})")

        # 7. Test Admin Reports API
        with client.session_transaction() as sess:
            sess['employee_code'] = 'ADMIN01'
            sess['role'] = 'admin'
        r = client.get('/api/admin/reports?dept=TDH')
        assert r.status_code == 200, f"Admin reports API failed: {r.status_code}"
        json_data = r.get_json()
        assert json_data.get('success') is True
        print("7. Admin Reports API: OK")

    print("\nALL SYSTEM INTEGRITY TESTS PASSED 100%!")

if __name__ == '__main__':
    test_routes()
